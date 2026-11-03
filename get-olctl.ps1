<#
    get-olctl.ps1 -- fetch the latest olctl release and install it. Needs no
    admin rights and no login.

        powershell -NoProfile -ExecutionPolicy Bypass -File get-olctl.ps1
        powershell -NoProfile -ExecutionPolicy Bypass -File get-olctl.ps1 -CheckOnly

    Finds the newest release tag through the repository's tags API, downloads
    that tag's source archive, and runs the install.ps1 inside it. The same
    script ships in the install folder, where `olctl update` runs it.

    Prints one JSON object (the last line of output). Exit 0 = installed, up
    to date, or checked; 1 = failed, with `code` and `error`.
#>

param(
    [string]$Version,
    [switch]$CheckOnly,
    [switch]$Force,
    [string]$Destination = (Join-Path $env:LOCALAPPDATA 'Programs\olctl'),
    [string]$SkillsRoot,
    [switch]$NoSkills,
    [switch]$NoPath,
    # Where releases come from. Anonymous read access is assumed. For an internal
    # GitLab, e.g.: .../api/v4/projects/<group>%2F<repo>/repository/tags,
    # '^olctl-(\d+\.\d+\.\d+)$', .../-/archive/{tag}/<repo>-{tag}.zip
    [string]$TagsUrl = 'https://api.github.com/repos/LeoStehlik/outlook-cli/tags?per_page=100',
    [string]$TagPattern = '^v(\d+\.\d+\.\d+)$',
    [string]$ArchiveUrl = 'https://github.com/LeoStehlik/outlook-cli/archive/refs/tags/{tag}.zip'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# For testing against another source (`olctl update` passes no URLs).
if ((-not $PSBoundParameters.ContainsKey('TagsUrl')) -and $env:OLCTL_TAGS_URL) { $TagsUrl = $env:OLCTL_TAGS_URL }
if ((-not $PSBoundParameters.ContainsKey('ArchiveUrl')) -and $env:OLCTL_ARCHIVE_URL) { $ArchiveUrl = $env:OLCTL_ARCHIVE_URL }

function Out-Result {
    param($Data, [int]$Rc)
    Write-Output (ConvertTo-Json -InputObject $Data -Depth 5 -Compress)
    exit $Rc
}

function Get-InstalledVersion {
    param([string]$Dir)
    $script = Join-Path $Dir 'olctl.ps1'
    if (-not (Test-Path -LiteralPath $script -PathType Leaf)) { return $null }
    $m = [regex]::Match([IO.File]::ReadAllText($script), "OLCTL_VERSION = '([^']+)'")
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

$result = [ordered]@{
    current = (Get-InstalledVersion $Destination); latest = $null; target = $null; tag = $null
    update_available = $false; installed = $false; destination = $Destination; install_output = @()
}
$stage = $null

try {
    # Windows PowerShell 5.1 does not offer TLS 1.2 by default.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    try {
        # 5.1 hands a JSON array back as one object; the pipeline unrolls it.
        $tags = @(Invoke-RestMethod -Uri $TagsUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'olctl' } | ForEach-Object { $_ })
    } catch {
        $result['ok'] = $false; $result['detail'] = $_.Exception.Message
        if ($null -ne $_.Exception.Response) {
            # The server answered: the repository is not readable without a login.
            $result['code'] = 'not_found'
            $result['error'] = 'the release list at ' + $TagsUrl + ' could not be read; the repository has to be public'
        } else {
            $result['code'] = 'no_network'
            $result['error'] = 'cannot reach ' + $TagsUrl + ' (offline, or blocked by a proxy?)'
        }
        Out-Result $result 1
    }
    $releases = New-Object System.Collections.Generic.List[object]
    foreach ($t in $tags) {
        $m = [regex]::Match([string]$t.name, $TagPattern)
        if ($m.Success) { [void]$releases.Add(@{ tag = [string]$t.name; version = [version]$m.Groups[1].Value }) }
    }
    if ($releases.Count -eq 0) {
        $result['ok'] = $false; $result['code'] = 'not_found'; $result['error'] = 'no release tags found at ' + $TagsUrl
        Out-Result $result 1
    }
    $newest = @($releases | Sort-Object { $_.version } -Descending)[0]
    $pick = $newest
    if (-not [string]::IsNullOrEmpty($Version)) {
        $pick = @($releases | Where-Object { $_.version -eq [version]$Version })[0]
        if ($null -eq $pick) {
            $result['ok'] = $false; $result['code'] = 'not_found'
            $result['error'] = 'no release ' + $Version + '; known: ' + [string]::Join(', ', @($releases | ForEach-Object { [string]$_.version }))
            Out-Result $result 1
        }
    }
    $result['latest'] = [string]$newest.version
    $result['target'] = [string]$pick.version
    $result['tag'] = $pick.tag
    $result['update_available'] = ($null -eq $result['current']) -or ([version]$result['current'] -lt $newest.version)

    if ($CheckOnly) { $result['ok'] = $true; Out-Result $result 0 }
    if ((-not $Force) -and ($result['current'] -eq [string]$pick.version)) {
        $result['ok'] = $true; $result['message'] = 'already up to date'
        Out-Result $result 0
    }

    $stage = Join-Path ([IO.Path]::GetTempPath()) ('olctl-get-' + [guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $stage)
    $zip = Join-Path $stage 'release.zip'
    $url = $ArchiveUrl.Replace('{tag}', $pick.tag).Replace('{version}', [string]$pick.version)
    try {
        Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing -Headers @{ 'User-Agent' = 'olctl' }
    } catch {
        $result['ok'] = $false; $result['code'] = 'no_network'
        $result['error'] = 'cannot download ' + $url; $result['detail'] = $_.Exception.Message
        Out-Result $result 1
    }
    $unpacked = Join-Path $stage 'src'
    Expand-Archive -LiteralPath $zip -DestinationPath $unpacked
    # A source archive holds one top folder; a package zip has the files at the top.
    $installer = @(Get-ChildItem -LiteralPath $unpacked -Filter 'install.ps1' -Recurse -Depth 1)[0]
    if ($null -eq $installer) {
        $result['ok'] = $false; $result['code'] = 'bad_install'
        $result['error'] = 'release ' + $pick.tag + ' has no install.ps1 (releases before 0.5.0 cannot be installed this way)'
        Out-Result $result 1
    }

    $installArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $installer.FullName, '-Destination', $Destination)
    if (-not [string]::IsNullOrEmpty($SkillsRoot)) { $installArgs += @('-SkillsRoot', $SkillsRoot) }
    if ($NoSkills) { $installArgs += '-NoSkills' }
    if ($NoPath) { $installArgs += '-NoPath' }
    $out = & powershell.exe @installArgs 2>&1 | ForEach-Object { [string]$_ }
    $rc = $LASTEXITCODE
    $result['install_output'] = @($out)
    if ($rc -ne 0) {
        $result['ok'] = $false; $result['code'] = 'install_failed'
        $result['error'] = 'install.ps1 from ' + $pick.tag + ' failed; see install_output'
        Out-Result $result 1
    }
    $result['installed'] = $true
    $result['previous'] = $result['current']
    $result['current'] = Get-InstalledVersion $Destination
    $result['ok'] = $true
    Out-Result $result 0
} catch {
    $result['ok'] = $false; $result['code'] = 'error'; $result['error'] = $_.Exception.Message
    Out-Result $result 1
} finally {
    if (($null -ne $stage) -and (Test-Path -LiteralPath $stage)) {
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
    }
}
