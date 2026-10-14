# In-memory stand-in for the PnP.PowerShell cmdlets the file manager uses.
# Lets the test suites run the real engines without a SharePoint tenant.
class Folder {
    [string]$Name; [string]$ServerRelativeUrl; [int]$ItemCount; [datetime]$TimeLastModified
}
class File {
    [string]$Name; [long]$Length; [string]$ServerRelativeUrl; [datetime]$TimeLastModified; [guid]$UniqueId
}

$global:FakeSite = 'https://contoso.sharepoint.com/sites/HR'
$global:FakeSitePath = '/sites/HR'
# key: server-relative path (lower) -> @{ Type='Folder'|'File'; Path; Bytes }
$global:Fake = @{}
$global:FakeFailList = @{}      # site-relative folders whose listing should fail
$global:FakeCopyJobError = $false
$global:FakeCalls = New-Object System.Collections.Generic.List[string]
$global:FakeRecycle = New-Object System.Collections.Generic.List[string]
$global:FakeUploads = New-Object System.Collections.Generic.List[string]

function Reset-Fake { $global:Fake = @{}; $global:FakeFailList = @{}; $global:FakeCalls.Clear(); $global:FakeRecycle.Clear(); $global:FakeUploads.Clear(); $global:FakeCopyJobError = $false }
function Add-FakeFolder { param([string]$Path) $global:Fake[$Path.ToLowerInvariant()] = @{ Type = 'Folder'; Path = $Path } }
function Add-FakeFile { param([string]$Path, [string]$Content) $global:Fake[$Path.ToLowerInvariant()] = @{ Type = 'File'; Path = $Path; Bytes = [Text.Encoding]::UTF8.GetBytes($Content) } }
function ConvertTo-FakeSrv { param([string]$SiteRel) if ($SiteRel.StartsWith('/')) { return $SiteRel.TrimEnd('/') } ; return ($global:FakeSitePath + '/' + $SiteRel).TrimEnd('/') }

function New-FakeObj {
    param($Node)
    $name = ($Node.Path -split '/')[-1]
    if ($Node.Type -eq 'Folder') {
        $f = [Folder]::new(); $f.Name = $name; $f.ServerRelativeUrl = $Node.Path; $f.TimeLastModified = [datetime]'2024-01-02'
        $f.ItemCount = @($global:Fake.Values | Where-Object { $_.Path -match ('^' + [regex]::Escape($Node.Path) + '/[^/]+$') }).Count
        return $f
    }
    $x = [File]::new(); $x.Name = $name; $x.Length = $Node.Bytes.Length; $x.ServerRelativeUrl = $Node.Path; $x.TimeLastModified = [datetime]::SpecifyKind([datetime]'2024-03-04 05:06:07', 'Utc'); $x.UniqueId = [guid]::NewGuid()
    return $x
}

function Connect-PnPOnline { param([string]$Url, [switch]$UseWebLogin) $global:FakeCalls.Add("Connect $Url") }
function Disconnect-PnPOnline { $global:FakeCalls.Add('Disconnect') }
function Get-PnPWeb { return @{ Title = 'Fake' } }
function Get-PnPFolderItem {
    param([string]$FolderSiteRelativeUrl)
    $srv = ConvertTo-FakeSrv $FolderSiteRelativeUrl
    $global:FakeCalls.Add("List $srv")
    if ($global:FakeFailList.ContainsKey($FolderSiteRelativeUrl.ToLowerInvariant())) { throw 'The operation has timed out (simulated throttle)' }
    if (-not $global:Fake.ContainsKey($srv.ToLowerInvariant())) { throw 'File Not Found.' }
    $pat = '^' + [regex]::Escape($srv) + '/[^/]+$'
    foreach ($n in @($global:Fake.Values | Where-Object { $_.Path -match $pat })) { New-FakeObj $n }
}
function Get-PnPFile {
    param([string]$Url, [string]$ServerRelativeUrl, [string]$Path, [string]$FileName, [switch]$AsFile, [switch]$Force)
    if ($AsFile) {
        $n = $global:Fake[$ServerRelativeUrl.ToLowerInvariant()]
        if (-not $n -or $n.Type -ne 'File') { throw 'File Not Found.' }
        [IO.File]::WriteAllBytes((Join-Path $Path $FileName), $n.Bytes)
        return
    }
    $srv = ConvertTo-FakeSrv $Url
    $n = $global:Fake[$srv.ToLowerInvariant()]
    if (-not $n -or $n.Type -ne 'File') { throw 'File Not Found.' }
    return (New-FakeObj $n)
}
function Remove-PnPFolder {
    param([string]$Name, [string]$Folder, [switch]$Force, [switch]$Recycle)
    $p = "$Folder/$Name"
    if (-not $global:Fake.ContainsKey($p.ToLowerInvariant())) { throw 'File Not Found.' }
    foreach ($k in @($global:Fake.Keys)) { if ($k -eq $p.ToLowerInvariant() -or $k.StartsWith($p.ToLowerInvariant() + '/')) { $global:Fake.Remove($k) } }
    $global:FakeRecycle.Add($p)
}
function Remove-PnPFile {
    param([string]$ServerRelativeUrl, [switch]$Force, [switch]$Recycle)
    if (-not $global:Fake.ContainsKey($ServerRelativeUrl.ToLowerInvariant())) { throw 'File Not Found.' }
    $global:Fake.Remove($ServerRelativeUrl.ToLowerInvariant()); $global:FakeRecycle.Add($ServerRelativeUrl)
}
function Copy-PnPFile {
    param([string]$SourceUrl, [string]$TargetUrl, [switch]$Force, [switch]$Overwrite)
    $global:FakeCalls.Add("Copy $SourceUrl -> $TargetUrl")
    if ($global:FakeCopyJobError) { return [pscustomobject]@{ Event = 'JobError'; Message = 'Simulated copy job failure' } }
    $leaf = ($SourceUrl -split '/')[-1]
    foreach ($n in @($global:Fake.Values | Where-Object { $_.Path -eq $SourceUrl -or $_.Path.StartsWith($SourceUrl + '/') })) {
        $np = $TargetUrl + '/' + $leaf + $n.Path.Substring($SourceUrl.Length)
        $global:Fake[$np.ToLowerInvariant()] = @{ Type = $n.Type; Path = $np; Bytes = $n.Bytes }
    }
    return [pscustomobject]@{ Event = 'JobEnd'; Message = 'ok' }
}
function Move-PnPFile { param([string]$SourceUrl, [string]$TargetUrl, [switch]$Force, [switch]$Overwrite) Copy-PnPFile -SourceUrl $SourceUrl -TargetUrl $TargetUrl; foreach ($k in @($global:Fake.Keys)) { if ($k -eq $SourceUrl.ToLowerInvariant() -or $k.StartsWith($SourceUrl.ToLowerInvariant() + '/')) { $global:Fake.Remove($k) } } }
function Resolve-PnPFolder {
    param([string]$SiteRelativePath)
    $acc = $global:FakeSitePath
    foreach ($seg in ($SiteRelativePath -split '/')) { if ($seg) { $acc = "$acc/$seg"; if (-not $global:Fake.ContainsKey($acc.ToLowerInvariant())) { Add-FakeFolder $acc } } }
}
function Add-PnPFile {
    param([string]$Path, [string]$Folder, [string]$NewFileName)
    $srv = (ConvertTo-FakeSrv $Folder) + '/' + $NewFileName
    $global:Fake[$srv.ToLowerInvariant()] = @{ Type = 'File'; Path = $srv; Bytes = [IO.File]::ReadAllBytes($Path) }
    $global:FakeUploads.Add($srv)
}
function Get-PnPList {
    $docs = [pscustomobject]@{ Title = 'Documents'; BaseType = 1; Hidden = $false; ItemCount = 3; RootFolder = [pscustomobject]@{ ServerRelativeUrl = "$($global:FakeSitePath)/Documents" } }
    $apps = [pscustomobject]@{ Title = 'Apps'; BaseType = 1; Hidden = $false; ItemCount = 1; RootFolder = [pscustomobject]@{ ServerRelativeUrl = "$($global:FakeSitePath)/Apps" } }
    $pages = [pscustomobject]@{ Title = 'Site Pages'; BaseType = 1; Hidden = $false; ItemCount = 1; RootFolder = [pscustomobject]@{ ServerRelativeUrl = "$($global:FakeSitePath)/SitePages" } }
    return @($apps, $docs, $pages)
}
