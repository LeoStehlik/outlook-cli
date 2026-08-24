# Regression harness for the argument layer. Runs anywhere PowerShell runs --
# it deliberately stops before COM, so a `no_outlook` result counts as "the
# arguments were accepted". The COM layer itself can only be tested on a
# machine with Outlook.
$ErrorActionPreference = 'Continue'
$script = Join-Path (Split-Path -Parent $PSScriptRoot) 'olctl.ps1'
$fails = 0

function Check {
    param([string]$Label, [string[]]$Argv, [string]$ExpectCode, [int]$ExpectRc = 1)
    $out = & pwsh -NoProfile -File $script @Argv 2>&1 | Out-String
    $rc = $LASTEXITCODE
    $ok = ($rc -eq $ExpectRc) -and ($out -match [regex]::Escape($ExpectCode))
    if ($ok) { Write-Output ("PASS  " + $Label) }
    else {
        $script:fails++
        Write-Output ("FAIL  " + $Label + "  rc=" + $rc + " out=" + $out.Trim())
    }
}

Check 'version'                  @('--version')                                   '"olctl_version"' 0
Check 'bad --since'              @('list','--folder','Inbox','--since','bogus')   'bad_argument'
Check 'list needs --folder'      @('list','--unread')                             'bad_argument'
Check 'get needs a ref'          @('get')                                         'bad_argument'
Check 'mark needs read|unread'   @('mark','A!B','sideways')                       'bad_argument'
Check 'body must be valid'       @('get','A!B','--body','wrong')                  'bad_argument'
Check 'unknown option'           @('list','--folder','Inbox','--nope')            'bad_argument'
Check 'unknown command'          @('frobnicate')                                  'bad_argument'
Check 'save-attachments --out'   @('save-attachments','A!B')                      'bad_argument'
Check 'move needs --to'          @('move','A!B')                                  'bad_argument'
Check 'draft-reply needs --text' @('draft-reply','A!B')                           'bad_argument'
Check 'draft-new needs --to'     @('draft-new','--text','hi')                     'bad_argument'
# These are well-formed, so they get as far as the COM attach.
Check 'doctor reaches COM'       @('doctor')                                      'no_outlook'
Check 'list reaches COM'         @('list','--folder','projects/inbox','--unread','--since','3d') 'no_outlook'
Check 'unicode folder accepted'  @('list','--folder','SA_AutomationAgent/Doručená pošta')                  'no_outlook'

# The payload must travel through the PIPELINE, not the console handle, or
# `.\olctl.ps1 ... | ConvertFrom-Json` captures nothing when the script is
# dot-invoked in-process. This regressed once; keep it tested.
function CheckCapture {
    param([string]$Label, [string[]]$Argv, [string]$Property, [string]$Expected)
    $captured = & $script $Argv 2>$null | Out-String
    $ok = $false
    $value = $null
    try {
        $obj = $captured | ConvertFrom-Json
        $value = $obj.$Property
        $ok = ([string]$value -eq $Expected)
    } catch { $ok = $false }
    if ($ok) { Write-Output ("PASS  " + $Label) }
    else {
        $script:fails++
        Write-Output ("FAIL  " + $Label + "  got=" + [string]$value + " captured=" + $captured.Trim())
    }
}

CheckCapture 'in-process pipe captures the payload' @('--version') 'olctl_version' '0.3.3'
CheckCapture 'in-process pipe captures errors too' @('list','--unread') 'code' 'bad_argument'
CheckCapture 'in-process pipe reaches COM'         @('doctor')        'code' 'no_outlook'

if ($fails -gt 0) { Write-Output ("`n" + $fails + " FAILED"); exit 1 }
Write-Output "`nall checks passed"
