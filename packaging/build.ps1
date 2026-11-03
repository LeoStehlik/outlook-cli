<#
    build.ps1 -- stage the deployment package and zip it.

        powershell -NoProfile -ExecutionPolicy Bypass -File packaging\build.ps1

    Produces dist\olctl-<version>\ (the folder to hand to IntuneWinAppUtil or
    ConfigMgr) and dist\olctl-<version>.zip. Runs the argument tests and the
    static contract tests first and refuses to build if either fails.
#>

param([switch]$SkipTests)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$m = [regex]::Match([IO.File]::ReadAllText((Join-Path $repo 'olctl.ps1')), "OLCTL_VERSION = '([^']+)'")
if (-not $m.Success) { throw 'cannot read the version from olctl.ps1' }
$version = $m.Groups[1].Value

if (-not $SkipTests) {
    foreach ($t in @('tests\cli-args.ps1', 'tests\contract.ps1')) {
        $out = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo $t) 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            Write-Output $out
            throw ($t + ' failed; not building')
        }
        Write-Output ($t + ': passed')
    }
}

$dist = Join-Path $repo 'dist'
$stage = Join-Path $dist ('olctl-' + $version)
if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
[void](New-Item -ItemType Directory -Path $stage -Force)

foreach ($p in @('olctl.ps1', 'olctl.cmd', 'guide', 'skills', 'install.ps1', 'uninstall.ps1', 'LICENSE')) {
    $from = Join-Path $repo $p
    if (Test-Path -LiteralPath $from) { Copy-Item -LiteralPath $from -Destination (Join-Path $stage $p) -Recurse -Force }
}
# The bash entry point must have LF endings whatever the checkout did to it.
$bash = [IO.File]::ReadAllText((Join-Path $repo 'olctl')).Replace("`r`n", "`n")
[IO.File]::WriteAllText((Join-Path $stage 'olctl'), $bash, (New-Object Text.UTF8Encoding($false)))

$zip = Join-Path $dist ('olctl-' + $version + '.zip')
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip

Write-Output ('Staged ' + $stage)
Write-Output ('Zipped ' + $zip)
Write-Output 'For Intune: IntuneWinAppUtil.exe -c <staged folder> -s install.ps1 -o <output folder>'
