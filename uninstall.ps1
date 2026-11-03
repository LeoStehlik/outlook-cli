<#
    uninstall.ps1 -- removes a per-user olctl install. Needs no admin rights.

        powershell -NoProfile -ExecutionPolicy Bypass -File uninstall.ps1

    Removes the install folder, its user PATH entry, the skills olctl
    installed (only folders carrying the .olctl-managed marker) and
    HKCU:\Software\olctl. Locations come from that registry key, so it works
    whether it is run from the deployment package or from the install folder.
    If the installer changed the user's execution policy, it is put back,
    unless it has been changed again since.
    The user's own data (%USERPROFILE%\.olctl: config, audit log, scan marks)
    is kept.
#>

param(
    [string]$Destination = $null,
    [string]$SkillsRoot = $null
)

$ErrorActionPreference = 'Stop'
$key = 'HKCU:\Software\olctl'
$Marker = '.olctl-install'
$SkillMarker = '.olctl-managed'
$rc = 0

try {
    $reg = $null
    if (Test-Path $key) { $reg = Get-ItemProperty -Path $key }
    if ([string]::IsNullOrEmpty($Destination)) {
        if (($null -ne $reg) -and $reg.InstallPath) { $Destination = [string]$reg.InstallPath }
        else { $Destination = Join-Path $env:LOCALAPPDATA 'Programs\olctl' }
    }
    if ([string]::IsNullOrEmpty($SkillsRoot)) {
        if (($null -ne $reg) -and $reg.SkillsRoot) { $SkillsRoot = [string]$reg.SkillsRoot }
        else { $SkillsRoot = Join-Path $env:USERPROFILE '.claude\skills' }
    }

    if (Test-Path -LiteralPath $SkillsRoot) {
        foreach ($dir in @(Get-ChildItem -LiteralPath $SkillsRoot -Directory)) {
            if (Test-Path -LiteralPath (Join-Path $dir.FullName $SkillMarker)) {
                Remove-Item -LiteralPath $dir.FullName -Recurse -Force
                Write-Output ('Removed skill ' + $dir.Name)
            }
        }
    }

    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    if ($null -ne $userPath) {
        $want = $Destination.TrimEnd('\')
        $parts = @($userPath.Split(';') | Where-Object { $_.Trim().Length -gt 0 })
        $kept = @($parts | Where-Object { $_.Trim().TrimEnd('\') -ine $want })
        if ($kept.Count -ne $parts.Count) {
            [Environment]::SetEnvironmentVariable('Path', [string]::Join(';', $kept), 'User')
            Write-Output 'Removed from the user PATH'
        }
    }

    if (Test-Path -LiteralPath $Destination) {
        if (Test-Path -LiteralPath (Join-Path $Destination $Marker)) {
            Remove-Item -LiteralPath $Destination -Recurse -Force
            Write-Output ('Removed ' + $Destination)
        } else {
            Write-Output ('Left ' + $Destination + ' alone: it was not created by the olctl installer')
        }
    }

    if (($null -ne $reg) -and ($null -ne $reg.PSObject.Properties['ExecutionPolicyBefore'])) {
        $before = [string]$reg.ExecutionPolicyBefore
        if ([string](Get-ExecutionPolicy -Scope CurrentUser) -eq 'RemoteSigned') {
            try { Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy $before -Force } catch { }
            Write-Output ('Restored the CurrentUser execution policy to ' + $before)
        }
    }

    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force
        Write-Output ('Removed ' + $key)
    }
    Write-Output 'Done. Your olctl settings and audit log in %USERPROFILE%\.olctl were kept.'
} catch {
    Write-Output ('FAILED: ' + $_.Exception.Message)
    $rc = 1
}
exit $rc
