<#
    Contract tests: the JSON shape of every command's response.

    The shapes below ARE the contract that skills, scripts and agents rely on.
    A key may be added in a minor version; removing or retyping one is a
    breaking change and needs a major version.

        powershell -File tests\contract.ps1           # examples\*.json only, no Outlook needed
        powershell -File tests\contract.ps1 -Live     # also real read-only runs against Outlook

    -Live only runs read-only commands (doctor --quick, list, get, thread, a
    batch of get) against -Folder (default: the user's inbox).
#>

param([switch]$Live, [string]$Folder = 'inbox')

$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
$script = Join-Path $repo 'olctl.ps1'
$fails = 0

# Types: string, int, bool, array, object, any; a trailing ? allows null.
$row = [ordered]@{
    ref = 'string'; subject = 'string?'; from_name = 'string?'; from_email = 'string?'
    received = 'string?'; unread = 'bool'; has_attachments = 'bool'; attachment_count = 'int'
    categories = 'array'; flag = 'string?'; importance = 'int'; size = 'int'
    conversation_id = 'string?'; folder = 'string?'; to = 'string?'
}
function Join-Shape { param($A, $B) $h = [ordered]@{}; foreach ($k in $A.Keys) { $h[$k] = $A[$k] }; foreach ($k in $B.Keys) { $h[$k] = $B[$k] }; return $h }

$Shapes = @{
    list   = @{ keys = [ordered]@{ folder = 'string'; filter = 'string?'; returned = 'int'; scanned = 'int'
                    skipped_non_mail_or_unreadable = 'int'; excluded = 'int'; truncated = 'bool'
                    truncated_by = 'string?'; next_cursor = 'string?'; items = 'array' }
                items = @{ items = $row } }
    count  = @{ keys = [ordered]@{ folder = 'string'; filter = 'string?'; count = 'int' } }
    get    = @{ keys = (Join-Shape $row ([ordered]@{ recipients = 'array'; attachments = 'array'; body_truncated = 'bool'
                    sent = 'string?'; reply_to = 'string?'; internet_message_id = 'string?'
                    sent_on_behalf_of = 'string?'; send_using_account = 'string?' })) }
    thread = @{ keys = [ordered]@{ ref = 'string'; conversation_topic = 'string?'; conversation_id = 'string?'
                    method = 'array'; count = 'int'; messages = 'array' }
                items = @{ messages = (Join-Shape $row ([ordered]@{ sent = 'string?'; match = 'string' })) } }
    doctor = @{ keys = [ordered]@{ outlook_version = 'string?'; profile = 'string?'; current_user = 'string?'
                    stores = 'array'; policy = 'object' } }
    folders = @{ keys = [ordered]@{ count = 'int'; folders = 'array' } }
    move   = @{ keys = [ordered]@{ subject = 'string?'; from = 'string?'; to = 'string?'; old_ref = 'string'
                    new_ref = 'string?'; dry_run = 'bool' } }
    mark   = @{ keys = [ordered]@{ subject = 'string?'; unread = 'bool'; ref = 'string'; dry_run = 'bool' } }
    flag   = @{ keys = [ordered]@{ subject = 'string?'; flag = 'string?'; ref = 'string'; dry_run = 'bool' } }
    categorize = @{ keys = [ordered]@{ subject = 'string?'; before = 'array'; after = 'array'; ref = 'string'; dry_run = 'bool' } }
    draft_reply = @{ keys = [ordered]@{ draft_ref = 'string?'; in_reply_to = 'string'; subject = 'string?'
                    from = 'string?'; from_mode = 'string?'; saved_to = 'string?'; sent = 'bool'; dry_run = 'bool' } }
    draft_new = @{ keys = [ordered]@{ draft_ref = 'string?'; subject = 'string?'; from = 'string?'
                    from_mode = 'string?'; saved_to = 'string?'; sent = 'bool'; dry_run = 'bool' } }
    save_attachments = @{ keys = [ordered]@{ ref = 'string'; subject = 'string?'; attachments = 'int'
                    saved = 'array'; skipped = 'array'; out_dir = 'string'; dry_run = 'bool' } }
    batch  = @{ keys = [ordered]@{ count = 'int'; succeeded = 'int'; failed = 'int'; not_run = 'int'
                    dry_run = 'bool'; results = 'array' } }
    update = @{ keys = [ordered]@{ current = 'string?'; latest = 'string'; target = 'string'; tag = 'string'
                    update_available = 'bool'; installed = 'bool'; destination = 'string'
                    install_output = 'array'; dry_run = 'bool' } }
    error  = @{ keys = [ordered]@{ ok = 'bool'; code = 'string'; error = 'string' } }
}

function Test-Type {
    param($Value, [string]$Type)
    $nullable = $Type.EndsWith('?')
    $t = $Type.TrimEnd('?')
    if ($null -eq $Value) { return ($nullable -or ($t -eq 'any')) }
    switch ($t) {
        'string' { return ($Value -is [string]) }
        'int'    { return (($Value -is [int]) -or ($Value -is [long]) -or ($Value -is [decimal])) }
        'bool'   { return ($Value -is [bool]) }
        'array'  { return ($Value -is [System.Array]) }
        'object' { return ($Value -is [System.Management.Automation.PSCustomObject]) }
        default  { return $true }
    }
}

function Test-Shape {
    param([string]$Label, $Obj, [string]$ShapeName)
    $shape = $Shapes[$ShapeName]
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($k in $shape.keys.Keys) {
        $p = $Obj.PSObject.Properties[$k]
        if ($null -eq $p) { [void]$problems.Add('missing ' + $k); continue }
        if (-not (Test-Type $p.Value $shape.keys[$k])) {
            [void]$problems.Add($k + ' should be ' + $shape.keys[$k] + ', got ' + $(if ($null -eq $p.Value) { 'null' } else { $p.Value.GetType().Name }))
        }
    }
    if ($shape.ContainsKey('items')) {
        foreach ($listKey in $shape.items.Keys) {
            $i = 0
            foreach ($it in @($Obj.$listKey)) {
                foreach ($k in $shape.items[$listKey].Keys) {
                    $p = $it.PSObject.Properties[$k]
                    if ($null -eq $p) { [void]$problems.Add($listKey + '[' + $i + '] missing ' + $k); continue }
                    if (-not (Test-Type $p.Value $shape.items[$listKey][$k])) {
                        [void]$problems.Add($listKey + '[' + $i + '].' + $k + ' should be ' + $shape.items[$listKey][$k])
                    }
                }
                $i++
            }
        }
    }
    if ($problems.Count -eq 0) { Write-Output ('PASS  ' + $Label) }
    else {
        $script:fails++
        Write-Output ('FAIL  ' + $Label + ': ' + [string]::Join('; ', $problems.ToArray()))
    }
}

# ------------------------------------------------------------- examples
$examples = @{
    'list.json' = 'list'; 'get.json' = 'get'; 'thread.json' = 'thread'; 'doctor.json' = 'doctor'
    'folders.json' = 'folders'; 'move.json' = 'move'; 'mark.json' = 'mark'; 'flag.json' = 'flag'
    'categorize.json' = 'categorize'; 'draft-reply.json' = 'draft_reply'; 'draft-new.json' = 'draft_new'
    'save-attachments.json' = 'save_attachments'; 'batch.json' = 'batch'; 'update.json' = 'update'; 'error-not-found.json' = 'error'
}
foreach ($f in ($examples.Keys | Sort-Object)) {
    $path = Join-Path (Join-Path $repo 'examples') $f
    if (-not (Test-Path -LiteralPath $path)) { $fails++; Write-Output ('FAIL  examples/' + $f + ': missing'); continue }
    $obj = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
    Test-Shape ('examples/' + $f) $obj $examples[$f]
}

# ------------------------------------------------------------------ live
if ($Live) {
    function Invoke-Live {
        param([string[]]$Argv)
        $t = & powershell -NoProfile -ExecutionPolicy Bypass -File $script @Argv | Out-String
        return ($t | ConvertFrom-Json)
    }
    Test-Shape 'live doctor --quick' (Invoke-Live @('doctor', '--quick')) 'doctor'
    $l = Invoke-Live @('list', '--folder', $Folder, '--since', '14d', '--limit', '3')
    Test-Shape 'live list' $l 'list'
    Test-Shape 'live list --count-only' (Invoke-Live @('list', '--folder', $Folder, '--since', '14d', '--count-only')) 'count'
    Test-Shape 'live error' (Invoke-Live @('list', '--folder', 'no-such-folder-olctl-test')) 'error'
    $first = @($l.items)[0]
    if ($null -ne $first) {
        Test-Shape 'live get' (Invoke-Live @('get', $first.ref, '--body', 'none')) 'get'
        Test-Shape 'live thread' (Invoke-Live @('thread', $first.ref, '--body', 'none')) 'thread'
        $bf = Join-Path ([IO.Path]::GetTempPath()) ('olctl-contract-' + [guid]::NewGuid().ToString('N') + '.jsonl')
        [IO.File]::WriteAllText($bf, (ConvertTo-Json -Compress @('get', $first.ref, '--body', 'none')))
        Test-Shape 'live batch' (Invoke-Live @('batch', '--file', $bf)) 'batch'
        Remove-Item -LiteralPath $bf -Force
    } else {
        Write-Output ('SKIP  live get/thread/batch: no mail in ' + $Folder + ' in the last 14 days')
    }
}

if ($fails -gt 0) { Write-Output ("`n" + $fails + ' FAILED'); exit 1 }
Write-Output "`nall contract checks passed"
