# Regression harness for the argument layer. Runs anywhere PowerShell runs --
# it deliberately stops before COM. A command that gets past argument checking
# reports no_outlook on a machine without Outlook, or whatever Outlook says on
# one with it; either way it is not bad_argument, and that is what is checked.
# The COM layer itself can only be tested on a machine with Outlook.
$ErrorActionPreference = 'Continue'
$script = Join-Path (Split-Path -Parent $PSScriptRoot) 'olctl.ps1'
$fails = 0

$shell = 'pwsh'
if (-not (Get-Command pwsh -ErrorAction SilentlyContinue)) { $shell = 'powershell' }

# Marks, presets and the audit log live under OLCTL_HOME; keep the run away
# from the real one.
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('olctl-test-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $tmp)
$env:OLCTL_HOME = $tmp
$cfg = '{"presets": {"noise": {"exclude_sender": ["noreply@"]}}}'
[IO.File]::WriteAllText((Join-Path $tmp 'config.json'), $cfg)

function Invoke-Olctl {
    param([string[]]$Argv)
    $out = & $shell -NoProfile -ExecutionPolicy Bypass -File $script @Argv 2>&1 | Out-String
    return @{ out = $out; rc = $LASTEXITCODE }
}

function Check {
    param([string]$Label, [string[]]$Argv, [string]$ExpectCode, [int]$ExpectRc = 1)
    $r = Invoke-Olctl $Argv
    $ok = ($r.rc -eq $ExpectRc) -and ($r.out -match [regex]::Escape($ExpectCode))
    if ($ok) { Write-Output ("PASS  " + $Label) }
    else {
        $script:fails++
        Write-Output ("FAIL  " + $Label + "  rc=" + $r.rc + " out=" + $r.out.Trim())
    }
}

# Well-formed: must get past the argument layer.
function CheckAccepted {
    param([string]$Label, [string[]]$Argv)
    $r = Invoke-Olctl $Argv
    if ($r.out -notmatch '"bad_argument"') { Write-Output ("PASS  " + $Label) }
    else {
        $script:fails++
        Write-Output ("FAIL  " + $Label + "  out=" + $r.out.Trim())
    }
}

$batchBad = Join-Path $tmp 'bad.jsonl'
[IO.File]::WriteAllText($batchBad, "[""mark"", ""A!B"", ""read""]`n{""not"": ""an array""}`n")
$batchNested = Join-Path $tmp 'nested.jsonl'
[IO.File]::WriteAllText($batchNested, "[""batch""]`n")
$batchArgs = Join-Path $tmp 'args.jsonl'
[IO.File]::WriteAllText($batchArgs, "[""mark"", ""A!B"", ""sideways""]`n")
$cursorOld = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"r":"2026-11-01T10:00:00","ref":"A!B","o":true,"t":null}'))

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
Check 'strip-quoted is text only'   @('get','A!B','--body','html','--strip-quoted')  'bad_argument'
Check 'thread needs a ref'          @('thread')                                      'bad_argument'
Check 'thread body text|none'       @('thread','A!B','--body','html')                'bad_argument'
Check 'draft-update needs a ref'    @('draft-update')                                'bad_argument'
Check 'draft-update needs a change' @('draft-update','A!B')                          'bad_argument'
Check 'whole-body needs --text'     @('draft-update','A!B','--whole-body','--subject','x') 'bad_argument'
Check 'cursor must be a cursor'     @('list','--folder','Inbox','--after-cursor','garbage') 'bad_argument'
Check 'cursor order must match'     @('list','--folder','Inbox','--after-cursor',$cursorOld) 'other order'
Check 'mark name is checked'        @('list','--folder','Inbox','--since-mark','no good')   'bad_argument'
Check 'unknown mark needs --since'  @('list','--folder','Inbox','--since-mark','intake')  'known_marks'
Check 'unknown preset'              @('list','--folder','Inbox','--preset','nope')   'known_presets'
Check 'batch rejects bad lines'     @('batch','--file',$batchBad)                    'nothing was run'
Check 'batch cannot nest'           @('batch','--file',$batchNested)                 'cannot be nested'
Check 'batch validates argv'        @('batch','--file',$batchArgs)                   'nothing was run'
Check 'batch file must exist'       @('batch','--file',(Join-Path $tmp 'missing.jsonl')) 'bad_argument'

# These are well-formed, so they get as far as the COM attach (or beyond).
CheckAccepted 'doctor reaches COM'        @('doctor','--quick')
CheckAccepted 'list reaches COM'          @('list','--folder','projects/inbox','--unread','--since','3d','--count-only')
CheckAccepted 'unicode folder accepted'   @('list','--folder','SA_AutomationAgent/Doručená pošta','--count-only')
CheckAccepted 'multi-folder list'         @('list','--folder','projects/inbox','--folder','projects/sent items','--since','1d','--count-only')
CheckAccepted 'body search + exclusions'  @('list','--folder','projects/inbox','--since','1d','--body-contains','x','--preset','noise','--exclude-subject','y','--count-only')
CheckAccepted 'unknown mark with --since' @('list','--folder','projects/inbox','--since-mark','intake','--since','1d','--count-only')

# --json-out writes the payload as UTF-8 and prints where it went.
$jo = Join-Path $tmp 'version.json'
$r = Invoke-Olctl @('--version','--json-out',$jo)
$written = $null
try { $written = ([IO.File]::ReadAllText($jo) | ConvertFrom-Json).olctl_version } catch { $written = $null }
if (($r.rc -eq 0) -and ($r.out -match 'json_out') -and ($null -ne $written)) { Write-Output 'PASS  --json-out writes the file' }
else { $fails++; Write-Output ('FAIL  --json-out writes the file  out=' + $r.out.Trim()) }

# guide is plain text, bundled next to the script, with the policy appended.
$r = Invoke-Olctl @('guide')
if (($r.rc -eq 0) -and ($r.out -match 'House rules') -and ($r.out -match 'On this machine')) { Write-Output 'PASS  guide prints the bundled guide' }
else { $fails++; Write-Output ('FAIL  guide prints the bundled guide  rc=' + $r.rc + ' out=' + $r.out.Substring(0, [Math]::Min(200, $r.out.Length))) }

# Policy. OLCTL_POLICY_PATH can only tighten the machine policy, so it is
# safe to drive these checks with it.
$polRO = Join-Path $tmp 'policy-ro.json'
[IO.File]::WriteAllText($polRO, '{"read_only": true}')
$polDis = Join-Path $tmp 'policy-disabled.json'
[IO.File]::WriteAllText($polDis, '{"disabled_commands": ["move", "rules-add"]}')
$polBad = Join-Path $tmp 'policy-bad.json'
[IO.File]::WriteAllText($polBad, '{ not json')
$batchMark = Join-Path $tmp 'mark.jsonl'
[IO.File]::WriteAllText($batchMark, "[""mark"", ""A!B"", ""read""]`n")

$env:OLCTL_POLICY_PATH = $polRO
Check 'read-only blocks mark'            @('mark','A!B','read')                         'disabled_by_policy'
Check 'read-only blocks batch lines'     @('batch','--file',$batchMark)                 'disabled_by_policy'
CheckAccepted 'read-only allows --dry-run' @('mark','A!B','read','--dry-run')
CheckAccepted 'read-only allows list'    @('list','--folder','inbox','--since','1d','--count-only')
$r = Invoke-Olctl @('guide')
if ($r.out -match 'READ-ONLY') { Write-Output 'PASS  guide shows read-only policy' }
else { $fails++; Write-Output 'FAIL  guide shows read-only policy' }
$env:OLCTL_POLICY_PATH = $polDis
Check 'disabled command blocked'         @('move','A!B','--to','inbox/x','--dry-run')   'disabled_by_policy'
CheckAccepted 'other commands still run' @('mark','A!B','read','--dry-run')
$env:OLCTL_POLICY_PATH = $polBad
Check 'unreadable policy fails closed'   @('list','--folder','inbox','--since','1d')    'bad_policy'
Check 'version still works'              @('--version')                                 '"olctl_version"' 0
Remove-Item Env:\OLCTL_POLICY_PATH

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

CheckCapture 'in-process pipe captures the payload' @('--version') 'olctl_version' '0.5.0'
CheckCapture 'in-process pipe captures errors too' @('list','--unread') 'code' 'bad_argument'

Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue

if ($fails -gt 0) { Write-Output ("`n" + $fails + " FAILED"); exit 1 }
Write-Output "`nall checks passed"
