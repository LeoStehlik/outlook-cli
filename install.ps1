<#
    install.ps1 -- per-user install of olctl. Needs no admin rights.

        powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1

    Copies olctl into %LOCALAPPDATA%\Programs\olctl, puts that folder on the
    user's PATH, installs the bundled Claude Code skills into
    %USERPROFILE%\.claude\skills, and records the version under
    HKCU:\Software\olctl for deployment tools to detect. Running it again
    updates in place.

    If the user's PowerShell execution policy would block olctl.ps1 (the
    Windows client default is Restricted), it sets CurrentUser to
    RemoteSigned, which needs no admin, and records the previous value so
    uninstall.ps1 can put it back. A policy set by Group Policy is left alone.
    -NoExecutionPolicy skips this. The user's own data (%USERPROFILE%\.olctl: config,
    audit log, scan marks) is never touched.

    Skill folders are only written if olctl created them (they carry a
    .olctl-managed marker), so a user's own skill of the same name is left
    alone and reported.

    Exit 0 = installed, 1 = failed (the reason is printed and logged).
#>

param(
    [string]$Destination = (Join-Path $env:LOCALAPPDATA 'Programs\olctl'),
    [string]$SkillsRoot = (Join-Path $env:USERPROFILE '.claude\skills'),
    [switch]$NoSkills,
    [switch]$NoPath,
    [switch]$NoExecutionPolicy
)

$ErrorActionPreference = 'Stop'
$script:Log = New-Object System.Collections.Generic.List[string]

function Write-Step {
    param([string]$Text)
    $line = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '  ' + $Text
    [void]$script:Log.Add($line)
    Write-Output $Text
}

# Files that make up an install. Anything else in the package (this script,
# skills\) is installer material, not runtime.
$Payload = @('olctl.ps1', 'olctl.cmd', 'olctl', 'guide', 'uninstall.ps1', 'get-olctl.ps1')
$Marker = '.olctl-install'
$SkillMarker = '.olctl-managed'

try {
    $src = $PSScriptRoot
    $srcScript = Join-Path $src 'olctl.ps1'
    if (-not (Test-Path -LiteralPath $srcScript -PathType Leaf)) {
        throw ('olctl.ps1 not found next to the installer in ' + $src)
    }
    $m = [regex]::Match([IO.File]::ReadAllText($srcScript), "OLCTL_VERSION = '([^']+)'")
    if (-not $m.Success) { throw 'cannot read the version from olctl.ps1' }
    $version = $m.Groups[1].Value
    Write-Step ('Installing olctl ' + $version + ' to ' + $Destination)

    # Never overwrite a folder olctl did not create.
    if (Test-Path -LiteralPath $Destination) {
        $hasMarker = Test-Path -LiteralPath (Join-Path $Destination $Marker)
        $isEmpty = @(Get-ChildItem -LiteralPath $Destination -Force).Count -eq 0
        if (-not ($hasMarker -or $isEmpty)) {
            throw ($Destination + ' exists and was not created by this installer; not touching it')
        }
    } else {
        [void](New-Item -ItemType Directory -Path $Destination -Force)
    }

    foreach ($p in $Payload) {
        $from = Join-Path $src $p
        if (-not (Test-Path -LiteralPath $from)) { throw ('package is missing ' + $p) }
        $to = Join-Path $Destination $p
        if (Test-Path -LiteralPath $to) { Remove-Item -LiteralPath $to -Recurse -Force }
        Copy-Item -LiteralPath $from -Destination $to -Recurse -Force
    }
    [IO.File]::WriteAllText((Join-Path $Destination $Marker), $version)
    # Files from a download carry the internet mark, which RemoteSigned refuses.
    Get-ChildItem -LiteralPath $Destination -Recurse -File | Unblock-File
    Write-Step ('Copied ' + [string]::Join(', ', $Payload))

    if (-not $NoPath) {
        $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
        if ($null -eq $userPath) { $userPath = '' }
        $parts = @($userPath.Split(';') | Where-Object { $_.Trim().Length -gt 0 })
        $want = $Destination.TrimEnd('\')
        $present = @($parts | Where-Object { $_.Trim().TrimEnd('\') -ieq $want }).Count -gt 0
        if ($present) {
            Write-Step 'Already on the user PATH'
        } else {
            $parts += $Destination
            [Environment]::SetEnvironmentVariable('Path', [string]::Join(';', $parts), 'User')
            Write-Step 'Added to the user PATH (new terminals pick it up)'
        }
    }

    $policyBefore = $null
    if (-not $NoExecutionPolicy) {
        # This process runs with -ExecutionPolicy Bypass, so look at the scopes
        # that apply to the user's own shells: Group Policy first, then user,
        # then machine; nothing set means the client default, Restricted.
        $effective = 'Restricted'
        $source = 'default'
        foreach ($scope in @('MachinePolicy', 'UserPolicy', 'CurrentUser', 'LocalMachine')) {
            $v = [string](Get-ExecutionPolicy -Scope $scope)
            if ($v -ne 'Undefined') { $effective = $v; $source = $scope; break }
        }
        if (@('Restricted', 'AllSigned') -notcontains $effective) {
            Write-Step ('Execution policy is ' + $effective + ' (' + $source + '); left as it is')
        } elseif (@('MachinePolicy', 'UserPolicy') -contains $source) {
            Write-Step ('Execution policy is ' + $effective + ', set by Group Policy; olctl.ps1 will not run in PowerShell itself (Git Bash and cmd still work)')
        } else {
            $policyBefore = [string](Get-ExecutionPolicy -Scope CurrentUser)
            # Saved even when it reports that a narrower scope (this process's
            # Bypass) overrides it, so check the stored value, not the error.
            try { Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force } catch { }
            if ([string](Get-ExecutionPolicy -Scope CurrentUser) -eq 'RemoteSigned') {
                Write-Step ('Set the CurrentUser execution policy to RemoteSigned (was ' + $policyBefore + ')')
            } else {
                $policyBefore = $null
                Write-Step 'Could not set the CurrentUser execution policy; olctl.ps1 may not run in PowerShell itself (Git Bash and cmd still work)'
            }
        }
    }

    if (-not $NoSkills) {
        $skillsSrc = Join-Path $src 'skills'
        if (Test-Path -LiteralPath $skillsSrc) {
            if (-not (Test-Path -LiteralPath $SkillsRoot)) {
                [void](New-Item -ItemType Directory -Path $SkillsRoot -Force)
            }
            foreach ($dir in @(Get-ChildItem -LiteralPath $skillsSrc -Directory)) {
                $target = Join-Path $SkillsRoot $dir.Name
                if ((Test-Path -LiteralPath $target) -and
                    (-not (Test-Path -LiteralPath (Join-Path $target $SkillMarker)))) {
                    Write-Step ('Skipped skill ' + $dir.Name + ': ' + $target + ' exists and is not managed by olctl')
                    continue
                }
                if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
                Copy-Item -LiteralPath $dir.FullName -Destination $target -Recurse -Force
                [IO.File]::WriteAllText((Join-Path $target $SkillMarker), $version)
                Write-Step ('Installed skill ' + $dir.Name)
            }
        }
    }

    $key = 'HKCU:\Software\olctl'
    if (-not (Test-Path $key)) { [void](New-Item -Path $key -Force) }
    Set-ItemProperty -Path $key -Name 'Version' -Value $version
    Set-ItemProperty -Path $key -Name 'InstallPath' -Value $Destination
    Set-ItemProperty -Path $key -Name 'SkillsRoot' -Value $SkillsRoot
    Set-ItemProperty -Path $key -Name 'InstalledAt' -Value ((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'))
    if ($null -ne $policyBefore) { Set-ItemProperty -Path $key -Name 'ExecutionPolicyBefore' -Value $policyBefore }
    Write-Step ('Recorded version ' + $version + ' under ' + $key)

    Write-Step 'Done. Check with: olctl doctor --quick'
    $rc = 0
} catch {
    Write-Step ('FAILED: ' + $_.Exception.Message)
    $rc = 1
}

try {
    if (Test-Path -LiteralPath $Destination) {
        [IO.File]::AppendAllLines((Join-Path $Destination 'install.log'), $script:Log)
    }
} catch { }
exit $rc
