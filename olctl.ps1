<#
    olctl.ps1 -- a local Outlook (Classic) CLI.

    Drives the Outlook desktop client on this machine over COM, as the logged-in
    user, using the mail profile Outlook already has. No Azure app registration,
    no admin consent, no app password, no network calls of its own.

        olctl.ps1 list --folder "projects/inbox" --unread --since 3d

    Prints exactly one JSON object on stdout. Exit 0 = ok, 1 = handled error.
    Nothing is ever sent and nothing is ever hard-deleted. Every mutation is
    appended to %USERPROFILE%\.olctl\audit.jsonl.

    Windows PowerShell 5.1 compatible: no ternary, no ??, no -AsHashtable,
    every ConvertTo-Json passes -Depth, and every JSON array is built as a
    List[object] then wrapped at the assignment site so a one-element array
    never collapses into a bare object.

    This file never calls Send() on anything.
#>

param([Parameter(ValueFromRemainingArguments = $true)][string[]]$CliArgs)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'

# ------------------------------------------------------------------ constants

$script:OL_FOLDER = @{
    'deleted items' = 3
    'outbox'        = 4
    'sent items'    = 5
    'inbox'         = 6
    'calendar'      = 9
    'contacts'      = 10
    'journal'       = 11
    'notes'         = 12
    'tasks'         = 13
    'drafts'        = 16
    'junk email'    = 23
    'junk'          = 23
    # No 'archive' entry on purpose. Two reasons. OlDefaultFolders 31 is documented as
    # olFolderArchive but this Outlook build returns "Quick Step Settings" for
    # it, so resolving the token by type would silently target the wrong folder.
    # A real Archive folder is found by its literal name like any other child.
    # And GetDefaultFolder CREATES a missing folder rather than failing, so
    # asking for a type a store lacks silently writes to the mailbox: calling it
    # with 31 once created a Quick Step Settings folder in four mailboxes here.
}

$script:OL_MAIL_ITEM = 43
$script:OL_FORMAT_PLAIN = 1

# Folder types that must never be a move destination without an override.
$script:OL_PROTECTED_TYPES = @(3, 4, 5, 23)

$script:PR_SMTP_ADDRESS = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
$script:PR_INTERNET_HEADERS = 'http://schemas.microsoft.com/mapi/proptag/0x007D001E'
$script:PR_INTERNET_MESSAGE_ID = 'http://schemas.microsoft.com/mapi/proptag/0x1035001E'
# Named string property (PS_PUBLIC_STRINGS) holding the text olctl last wrote
# into a draft, so draft-update can swap that text and keep the quoted original.
$script:OLCTL_DRAFT_TEXT = 'http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/olctl_text'

$script:RECIP_TYPE = @{ 1 = 'to'; 2 = 'cc'; 3 = 'bcc' }

# OlExchangeStoreType. Reported by name when recognised, raw int otherwise --
# the numbering has shifted between Outlook versions, so never trust it blindly.
$script:EXCHANGE_STORE_TYPE = @{
    0 = 'primary_exchange_mailbox'
    1 = 'exchange_mailbox'
    2 = 'exchange_public_folder'
    3 = 'not_exchange'
    4 = 'additional_exchange_mailbox'
}

$script:OlErrorPayload = $null
$script:ProtectedIdsCache = $null

# -------------------------------------------------------------------- plumbing

function New-OlErrorPayload {
    param(
        [string]$Code = 'error',
        [string]$Message = 'unspecified error',
        $Detail = $null,
        $Extra = $null
    )
    $h = @{}
    $h['__olctl_error'] = $true
    $h['code'] = $Code
    $h['error'] = $Message
    $h['detail'] = $Detail
    if ($null -ne $Extra) {
        foreach ($k in @($Extra.Keys)) { $h[$k] = $Extra[$k] }
    }
    return $h
}

# Raise a structured, reportable error. The payload is stashed in script scope
# as well as thrown, because relying on ErrorRecord.TargetObject alone is
# fragile across PowerShell hosts.
function Invoke-OlFail {
    param(
        [string]$Code = 'error',
        [string]$Message = 'unspecified error',
        $Detail = $null,
        $Extra = $null
    )
    $payload = New-OlErrorPayload -Code $Code -Message $Message -Detail $Detail -Extra $Extra
    $script:OlErrorPayload = $payload
    throw $payload
}

# Read a scalar property that may not exist / may throw. Internal variable
# names are deliberately obscure so a caller's scriptblock cannot be shadowed.
function Get-Safe {
    param($Fn, $Default = $null)
    try {
        $__olv = & $Fn
        if ($null -eq $__olv) { return $Default }
        return $__olv
    } catch {
        return $Default
    }
}

function Get-OlArg {
    param($Obj, [string]$Name, $Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $null
    try { $p = $Obj.PSObject.Properties[$Name] } catch { $p = $null }
    if ($null -eq $p) { return $Default }
    if ($null -eq $p.Value) { return $Default }
    return $p.Value
}

function Test-OlArgPresent {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $false }
    $p = $null
    try { $p = $Obj.PSObject.Properties[$Name] } catch { $p = $null }
    if ($null -eq $p) { return $false }
    if ($null -eq $p.Value) { return $false }
    return $true
}

# Always yields an [object[]]; null and missing become an empty array rather
# than an array containing one $null.
# NOTE: returns exactly ONE [object[]] (the `, @(...)` guard stops the pipeline
# from unrolling it). Assign the result directly -- never write @(Helper ...),
# because @( ) around a *command* adds another collection level and you end up
# with a nested array in the JSON.
function Get-OlArgArray {
    param($Obj, [string]$Name)
    $v = $null
    if ($null -ne $Obj) {
        $p = $null
        try { $p = $Obj.PSObject.Properties[$Name] } catch { $p = $null }
        if ($null -ne $p) { $v = $p.Value }
    }
    $out = New-Object System.Collections.Generic.List[object]
    if ($null -ne $v) {
        if (($v -is [System.Collections.IEnumerable]) -and -not ($v -is [string])) {
            foreach ($e in $v) { [void]$out.Add($e) }
        } else {
            [void]$out.Add($v)
        }
    }
    return , @($out.ToArray())
}

function Get-OlBool {
    param($Obj, [string]$Name, [bool]$Default = $false)
    $v = Get-OlArg $Obj $Name $null
    if ($null -eq $v) { return $Default }
    try { return [bool]$v } catch { return $Default }
}

function Get-OlInt {
    param($Obj, [string]$Name, [int]$Default = 0)
    $v = Get-OlArg $Obj $Name $null
    if ($null -eq $v) { return $Default }
    try { return [int]$v } catch { return $Default }
}

function Get-OlString {
    param($Obj, [string]$Name, $Default = $null)
    $v = Get-OlArg $Obj $Name $null
    if ($null -eq $v) { return $Default }
    try { return [string]$v } catch { return $Default }
}

function Get-OlIso {
    param($Value)
    if ($null -eq $Value) { return $null }
    try {
        $d = [datetime]$Value
        return $d.ToString('yyyy-MM-dd\THH:mm:ss')
    } catch {
        try { return [string]$Value } catch { return $null }
    }
}

function Get-OlTruncated {
    param($Text, [int]$Max)
    if ($null -eq $Text) { return '' }
    $s = [string]$Text
    if ($Max -le 0) { return '' }
    if ($s.Length -le $Max) { return $s }
    return $s.Substring(0, $Max)
}

# A position in list order: received second plus ref as a tie-break, so the
# order is total and does not depend on how Outlook orders equal timestamps.
# Cursors point at a position; ordinal comparison keeps it culture-proof.
function Compare-OlPos {
    param($A, $B)
    $c = [string]::CompareOrdinal([string]$A['r'], [string]$B['r'])
    if ($c -ne 0) { return $c }
    return [string]::CompareOrdinal([string]$A['ref'], [string]$B['ref'])
}

# True when $Pos comes after $Mark in the walk order (newest-first unless
# $OldestFirst).
function Test-OlPosAfter {
    param($Pos, $Mark, [bool]$OldestFirst)
    $c = Compare-OlPos $Pos $Mark
    if ($OldestFirst) { return ($c -gt 0) }
    return ($c -lt 0)
}

function Get-OlPosComparison {
    param([bool]$OldestFirst)
    if ($OldestFirst) {
        return [Comparison[object]] { param($x, $y) Compare-OlPos $x $y }
    }
    return [Comparison[object]] { param($x, $y) Compare-OlPos $y $x }
}

function ConvertTo-OlCursor {
    param($Pos, [bool]$OldestFirst, [string]$Started)
    $h = @{ r = [string]$Pos['r']; ref = [string]$Pos['ref']; o = [bool]$OldestFirst; t = $Started }
    $json = ConvertTo-Json -InputObject $h -Compress
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
}

# Cut the quoted history off a plain-text reply. Matches the separators Outlook
# and common clients write in English, German and Czech. Non-ASCII is spelled
# as \u escapes: this file has no BOM, so PowerShell 5.1 reads it as ANSI.
$script:QUOTE_START = @(
    '^\s*-{2,}\s*(Original Message|Urspr\u00fcngliche Nachricht|Originalnachricht|P\u016fvodn\u00ed zpr\u00e1va)\s*-{2,}',
    '^\s*(On|Am)\s.+(wrote|schrieb)\s*:?\s*$'
)
$script:QUOTE_HEADER = '^\s*\*?(From|Von|Od)\s*:\*?\s*\S'
$script:QUOTE_HEADER_NEXT = '^\s*\*?(Sent|Gesendet|Odesl\u00e1no|Date|Datum)\s*:'

function Get-OlStrippedBody {
    param([string]$Text)
    $res = @{ text = ''; removed = 0 }
    if ([string]::IsNullOrEmpty($Text)) { return $res }
    $lines = [regex]::Split($Text, '\r?\n')
    $cut = -1
    for ($i = 0; $i -lt $lines.Length -and $cut -lt 0; $i++) {
        $ln = [string]$lines[$i]
        foreach ($p in $script:QUOTE_START) {
            if ([regex]::IsMatch($ln, $p)) { $cut = $i; break }
        }
        if ($cut -ge 0) { break }
        # Outlook: a ____ rule, then From:/Sent: lines. Or the header block
        # alone. Either way From: must be followed closely by Sent:/Date:.
        $hdrAt = -1
        if ([regex]::IsMatch($ln, '^\s*_{10,}\s*$')) {
            for ($j = $i + 1; $j -le ($i + 2) -and $j -lt $lines.Length; $j++) {
                if ([regex]::IsMatch([string]$lines[$j], $script:QUOTE_HEADER)) { $hdrAt = $j; break }
            }
        } elseif ([regex]::IsMatch($ln, $script:QUOTE_HEADER)) {
            $hdrAt = $i
        }
        if ($hdrAt -ge 0) {
            for ($j = $hdrAt + 1; $j -le ($hdrAt + 4) -and $j -lt $lines.Length; $j++) {
                if ([regex]::IsMatch([string]$lines[$j], $script:QUOTE_HEADER_NEXT)) { $cut = $i; break }
            }
        }
        if (($cut -lt 0) -and $ln.StartsWith('>') -and (($i + 1) -lt $lines.Length) -and ([string]$lines[$i + 1]).StartsWith('>')) {
            $cut = $i
        }
    }
    if ($cut -lt 0) {
        $res['text'] = $Text
        return $res
    }
    $kept = [string]::Join("`r`n", $lines[0..([Math]::Max(0, $cut - 1))])
    if ($cut -eq 0) { $kept = '' }
    $kept = $kept.TrimEnd()
    $res['text'] = $kept
    $res['removed'] = $Text.Length - $kept.Length
    return $res
}

# --------------------------------------------------------------- COM attach

function Connect-Outlook {
    $app = $null
    try {
        $app = New-Object -ComObject Outlook.Application
    } catch {
        Invoke-OlFail -Code 'no_outlook' -Detail $_.Exception.Message -Message ('could not start or attach to Outlook Classic over COM. ' +
            'Check that classic Outlook (outlook.exe) is installed and has a mail profile. ' +
            'The New Outlook app does not expose a COM interface.')
    }
    $ns = $null
    try {
        $ns = $app.GetNamespace('MAPI')
    } catch {
        Invoke-OlFail -Code 'no_outlook' -Detail $_.Exception.Message -Message ('Outlook started but the MAPI namespace could not be opened. ' +
            'Classic Outlook must be installed with a working mail profile; ' +
            'the New Outlook app does not expose a COM interface.')
    }
    try { $ns.Logon('', '', $false, $false) } catch { }
    $pair = @{}
    $pair['app'] = $app
    $pair['ns'] = $ns
    return $pair
}

# ------------------------------------------------------------ folder helpers

# NOTE: returns exactly ONE [object[]] (the `, @(...)` guard stops the pipeline
# from unrolling it). Assign the result directly -- never write @(Helper ...),
# because @( ) around a *command* adds another collection level and you end up
# with a nested array in the JSON.
function Split-OlPath {
    param([string]$Path)
    $out = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrEmpty($Path)) { return , @($out.ToArray()) }
    $buf = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $Path.Length) {
        $ch = $Path[$i]
        if (($ch -eq '\') -and (($i + 1) -lt $Path.Length) -and (($Path[$i + 1] -eq '/') -or ($Path[$i + 1] -eq '\'))) {
            # \/ and \\ escape a literal separator
            [void]$buf.Append($Path[$i + 1])
            $i = $i + 2
            continue
        }
        if (($ch -eq '/') -or ($ch -eq '\')) {
            if ($buf.Length -gt 0) { [void]$out.Add($buf.ToString()) }
            [void]$buf.Clear()
        } else {
            [void]$buf.Append($ch)
        }
        $i = $i + 1
    }
    if ($buf.Length -gt 0) { [void]$out.Add($buf.ToString()) }
    return , @($out.ToArray())
}

# NOTE: returns exactly ONE [object[]] (the `, @(...)` guard stops the pipeline
# from unrolling it). Assign the result directly -- never write @(Helper ...),
# because @( ) around a *command* adds another collection level and you end up
# with a nested array in the JSON.
function Get-OlChildNames {
    param($Folder)
    $names = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Folder) { return , @($names.ToArray()) }
    $n = 0
    try { $n = [int]$Folder.Folders.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $nm = $null
        try { $nm = [string]$Folder.Folders.Item($i).Name } catch { $nm = $null }
        if ($null -ne $nm) { [void]$names.Add($nm) }
    }
    return , @($names.ToArray())
}

function Find-OlChild {
    param($Folder, [string]$Name)
    if ($null -eq $Folder) { return $null }
    if ($null -eq $Name) { return $null }
    $want = $Name.Trim().ToLower()
    $n = 0
    try { $n = [int]$Folder.Folders.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $sub = $null
        try { $sub = $Folder.Folders.Item($i) } catch { $sub = $null }
        if ($null -eq $sub) { continue }
        $nm = $null
        try { $nm = [string]$sub.Name } catch { $nm = $null }
        if ($null -eq $nm) { continue }
        if ($nm.Trim().ToLower() -eq $want) { return $sub }
    }
    return $null
}

# NOTE: returns exactly ONE [object[]] (the `, @(...)` guard stops the pipeline
# from unrolling it). Assign the result directly -- never write @(Helper ...),
# because @( ) around a *command* adds another collection level and you end up
# with a nested array in the JSON.
function Get-OlStoreNames {
    param($Ns)
    $names = New-Object System.Collections.Generic.List[object]
    $n = 0
    try { $n = [int]$Ns.Folders.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $nm = $null
        try { $nm = [string]$Ns.Folders.Item($i).Name } catch { $nm = $null }
        if ($null -ne $nm) { [void]$names.Add($nm) }
    }
    return , @($names.ToArray())
}

function Get-OlStoreDefaultFolder {
    param($Store, [int]$Code)
    if ($null -eq $Store) { return $null }
    try { return $Store.GetDefaultFolder($Code) } catch { return $null }
}

function Get-OlFolderPath {
    param($Folder)
    if ($null -eq $Folder) { return $null }
    try {
        $p = [string]$Folder.FolderPath
        if (-not [string]::IsNullOrEmpty($p)) {
            return ($p.TrimStart([char[]]@('\')) -replace '\\', '/')
        }
    } catch { }
    try { return [string]$Folder.Name } catch { return $null }
}

<#
    Resolve a folder path to a COM folder object.

    With -Shared the path is relative to that mailbox's Inbox (an empty path is
    the Inbox itself). Otherwise the first segment may be a store/mailbox name,
    a well-known English folder token, or a top-level folder of the default
    store.
#>
function Resolve-OlFolder {
    param($Ns, [string]$Path, $Shared)

    $parts = New-Object System.Collections.Generic.List[object]
    foreach ($p in (Split-OlPath $Path)) { [void]$parts.Add([string]$p) }

    $cur = $null

    if (-not [string]::IsNullOrEmpty($Shared)) {
        $rec = $null
        try {
            $rec = $Ns.CreateRecipient($Shared)
            [void]$rec.Resolve()
        } catch {
            Invoke-OlFail -Code 'unresolved_mailbox' -Detail $_.Exception.Message `
                -Message ("cannot resolve mailbox '" + $Shared + "' in the address book")
        }
        $resolved = $false
        try { $resolved = [bool]$rec.Resolved } catch { $resolved = $false }
        if (-not $resolved) {
            Invoke-OlFail -Code 'unresolved_mailbox' `
                -Message ("cannot resolve mailbox '" + $Shared + "' in the address book")
        }
        try {
            $cur = $Ns.GetSharedDefaultFolder($rec, $script:OL_FOLDER['inbox'])
        } catch {
            Invoke-OlFail -Code 'no_access' -Detail $_.Exception.Message `
                -Message ("no permission to open the Inbox of '" + $Shared + "', or the mailbox is not reachable from this profile")
        }
        if ($null -eq $cur) {
            Invoke-OlFail -Code 'no_access' `
                -Message ("no permission to open the Inbox of '" + $Shared + "', or the mailbox is not reachable from this profile")
        }
        if (($parts.Count -gt 0) -and (([string]$parts[0]).Trim().ToLower() -eq 'inbox')) {
            $parts.RemoveAt(0)
        }
    } else {
        if ($parts.Count -eq 0) {
            Invoke-OlFail -Code 'bad_argument' -Message 'a folder path is required'
        }
        $head = [string]$parts[0]
        $parts.RemoveAt(0)
        $headKey = $head.Trim().ToLower()
        $storeRoot = $null

        $n = 0
        try { $n = [int]$Ns.Folders.Count } catch { $n = 0 }
        for ($i = 1; $i -le $n; $i++) {
            $st = $null
            try { $st = $Ns.Folders.Item($i) } catch { $st = $null }
            if ($null -eq $st) { continue }
            $nm = $null
            try { $nm = [string]$st.Name } catch { $nm = $null }
            if ($null -eq $nm) { continue }
            if ($nm.Trim().ToLower() -eq $headKey) {
                $cur = $st
                $storeRoot = $st
                break
            }
        }

        if (($null -eq $cur) -and $script:OL_FOLDER.ContainsKey($headKey)) {
            # Unqualified well-known token: the default store is what is meant.
            try { $cur = $Ns.GetDefaultFolder($script:OL_FOLDER[$headKey]) } catch { $cur = $null }
        }

        if ($null -eq $cur) {
            $root = $null
            try { $root = $Ns.GetDefaultFolder($script:OL_FOLDER['inbox']).Parent } catch { $root = $null }
            if ($null -ne $root) { $cur = Find-OlChild $root $head }
        }

        if ($null -eq $cur) {
            $stores = Get-OlStoreNames $Ns
            $tops = New-Object System.Collections.Generic.List[object]
            try {
                $droot = $Ns.GetDefaultFolder($script:OL_FOLDER['inbox']).Parent
                foreach ($c in (Get-OlChildNames $droot)) { [void]$tops.Add($c) }
            } catch { }
            $children = New-Object System.Collections.Generic.List[object]
            foreach ($s in $stores) { [void]$children.Add($s) }
            foreach ($t in $tops) { if (-not $children.Contains($t)) { [void]$children.Add($t) } }
            $wk = New-Object System.Collections.Generic.List[object]
            foreach ($k in (@($script:OL_FOLDER.Keys) | Sort-Object)) { [void]$wk.Add($k) }
            $extra = @{}
            $extra['children'] = @($children.ToArray())
            $extra['available_stores'] = @($stores)
            $extra['well_known'] = @($wk.ToArray())
            Invoke-OlFail -Code 'not_found' -Extra $extra `
                -Message ("no store or folder named '" + $head + "'")
        }

        # CRITICAL: a well-known token directly under a NAMED store must resolve
        # inside THAT store. Falling through to Namespace.GetDefaultFolder here
        # would silently hand back the default mailbox's folder -- i.e. act on
        # the wrong mailbox -- which matters doubly in a profile where folder
        # names are localised (Inbox / Posteingang / Doruc. posta) and a literal
        # child-name match therefore cannot succeed.
        if (($parts.Count -gt 0) -and ($null -ne $storeRoot)) {
            $seg = [string]$parts[0]
            $token = $seg.Trim().ToLower()
            if (($null -eq (Find-OlChild $cur $seg)) -and $script:OL_FOLDER.ContainsKey($token)) {
                $storeObj = Get-Safe { $storeRoot.Store } $null
                $viaStore = Get-OlStoreDefaultFolder $storeObj $script:OL_FOLDER[$token]
                if ($null -eq $viaStore) {
                    $storeChildren = Get-OlChildNames $cur
                    $extra = @{}
                    $extra['children'] = @($storeChildren)
                    Invoke-OlFail -Code 'not_found' -Extra $extra `
                        -Message ("store '" + $head + "' has no default folder for '" + $seg + "'")
                }
                $cur = $viaStore
                $parts.RemoveAt(0)
            }
        }
    }

    foreach ($part in $parts) {
        $nxt = Find-OlChild $cur ([string]$part)
        if ($null -eq $nxt) {
            $segChildren = Get-OlChildNames $cur
            $extra = @{}
            $extra['children'] = @($segChildren)
            $parentName = Get-Safe { $cur.Name } '?'
            Invoke-OlFail -Code 'not_found' -Extra $extra `
                -Message ("folder '" + [string]$part + "' not found under '" + $parentName + "'")
        }
        $cur = $nxt
    }

    return $cur
}

# -------------------------------------------------------------- item helpers

function Get-OlRef {
    param($Item)
    $storeId = ''
    try { $storeId = [string]$Item.Parent.StoreID } catch { $storeId = '' }
    if ($null -eq $storeId) { $storeId = '' }
    $entryId = ''
    try { $entryId = [string]$Item.EntryID } catch { $entryId = '' }
    if ($null -eq $entryId) { $entryId = '' }
    return ($entryId + '!' + $storeId)
}

function Get-OlItem {
    param($Ns, [string]$Ref)
    if ([string]::IsNullOrEmpty($Ref)) {
        Invoke-OlFail -Code 'bad_argument' -Message 'a ref is required, in the form <EntryID>!<StoreID>'
    }
    $entryId = $Ref
    $storeId = ''
    $idx = $Ref.IndexOf('!')
    if ($idx -ge 0) {
        $entryId = $Ref.Substring(0, $idx)
        $storeId = $Ref.Substring($idx + 1)
    }
    $item = $null
    try {
        if ([string]::IsNullOrEmpty($storeId)) {
            $item = $Ns.GetItemFromID($entryId)
        } else {
            $item = $Ns.GetItemFromID($entryId, $storeId)
        }
    } catch {
        Invoke-OlFail -Code 'not_found' -Detail $_.Exception.Message -Message ('no item with that ref. Note that a ref changes when the item is ' +
            'moved between folders, so re-run list/get after a move.')
    }
    if ($null -eq $item) {
        Invoke-OlFail -Code 'not_found' -Message ('no item with that ref. Note that a ref changes when the item is ' +
            'moved between folders, so re-run list/get after a move.')
    }
    return $item
}

function Get-OlSenderSmtp {
    param($Item)
    $kind = ''
    try { $kind = [string]$Item.SenderEmailType } catch { $kind = '' }
    if ($kind -eq 'EX') {
        try {
            $v = $Item.Sender.GetExchangeUser().PrimarySmtpAddress
            if (-not [string]::IsNullOrEmpty([string]$v)) { return [string]$v }
        } catch { }
        try {
            $v = $Item.PropertyAccessor.GetProperty($script:PR_SMTP_ADDRESS)
            if (-not [string]::IsNullOrEmpty([string]$v)) { return [string]$v }
        } catch { }
    }
    try {
        $v = $Item.SenderEmailAddress
        if ($null -eq $v) { return $null }
        return [string]$v
    } catch {
        return $null
    }
}

function Get-OlRecipientSmtp {
    param($Recipient)
    try {
        $addr = [string]$Recipient.Address
        if ((-not [string]::IsNullOrEmpty($addr)) -and (-not $addr.StartsWith('/'))) { return $addr }
    } catch { }
    try {
        $v = $Recipient.PropertyAccessor.GetProperty($script:PR_SMTP_ADDRESS)
        if ($null -eq $v) { return $null }
        return [string]$v
    } catch {
        return $null
    }
}

function Get-OlFlagState {
    param($Item)
    $status = 0
    try { $status = [int](Get-Safe { $Item.FlagStatus } 0) } catch { $status = 0 }
    if ($status -eq 2) { return 'flagged' }
    if ($status -eq 1) { return 'completed' }
    return $null
}

# NOTE: returns exactly ONE [object[]] (the `, @(...)` guard stops the pipeline
# from unrolling it). Assign the result directly -- never write @(Helper ...),
# because @( ) around a *command* adds another collection level and you end up
# with a nested array in the JSON.
function Get-OlCategoryList {
    param($Item)
    $cats = New-Object System.Collections.Generic.List[object]
    $raw = Get-Safe { $Item.Categories } ''
    if ($null -ne $raw) {
        foreach ($c in ([string]$raw).Split(';')) {
            $t = ([string]$c).Trim()
            if ($t.Length -gt 0) { [void]$cats.Add($t) }
        }
    }
    return , @($cats.ToArray())
}

<#
    EntryIDs of the protected default folders of *every* store.

    Name matching is not good enough: one profile can hold Deleted Items,
    Geloeschte Elemente, Geloeschte Objekte and Odstranena posta at once.
    Resolving each store's own defaults makes the guard language-independent.
#>
function Get-OlProtectedFolderIds {
    param($Ns)
    if ($null -ne $script:ProtectedIdsCache) { return $script:ProtectedIdsCache }
    $ids = @{}
    try {
        $stores = $Ns.Stores
        $n = 0
        try { $n = [int]$stores.Count } catch { $n = 0 }
        for ($i = 1; $i -le $n; $i++) {
            $store = $null
            try { $store = $stores.Item($i) } catch { $store = $null }
            if ($null -eq $store) { continue }
            $label = [string](Get-Safe { $store.DisplayName } '?')
            foreach ($code in $script:OL_PROTECTED_TYPES) {
                $folder = Get-OlStoreDefaultFolder $store ([int]$code)
                if ($null -eq $folder) { continue }
                $entry = $null
                try { $entry = [string]$folder.EntryID } catch { $entry = $null }
                if ([string]::IsNullOrEmpty($entry)) { continue }
                $fname = [string](Get-Safe { $folder.Name } '?')
                $ids[$entry] = ($label + '/' + $fname)
            }
        }
    } catch { }
    $script:ProtectedIdsCache = $ids
    return $ids
}

# ------------------------------------------------------------ item projection

function Get-OlRow {
    param($Item, [int]$PreviewChars = 0)

    $row = @{}
    $row['ref'] = Get-Safe { Get-OlRef -Item $Item } $null
    $row['subject'] = Get-Safe { $Item.Subject } ''
    $row['from_name'] = Get-Safe { $Item.SenderName } $null
    $row['from_email'] = Get-OlSenderSmtp -Item $Item
    $row['received'] = Get-OlIso (Get-Safe { $Item.ReceivedTime } $null)
    $row['unread'] = [bool](Get-Safe { $Item.UnRead } $false)

    $attCount = 0
    try { $attCount = [int](Get-Safe { $Item.Attachments.Count } 0) } catch { $attCount = 0 }
    $row['has_attachments'] = ($attCount -gt 0)
    $row['attachment_count'] = $attCount

    $rowCats = Get-OlCategoryList -Item $Item
    $row['categories'] = @($rowCats)
    $row['flag'] = Get-OlFlagState -Item $Item

    $imp = 1
    try { $imp = [int](Get-Safe { $Item.Importance } 1) } catch { $imp = 1 }
    $row['importance'] = $imp

    $sz = 0
    try { $sz = [int](Get-Safe { $Item.Size } 0) } catch { $sz = 0 }
    $row['size'] = $sz

    $row['conversation_id'] = Get-Safe { $Item.ConversationID } $null
    $row['folder'] = Get-Safe { Get-OlFolderPath -Folder $Item.Parent } $null
    $row['to'] = Get-Safe { $Item.To } $null

    if ($PreviewChars -gt 0) {
        $body = [string](Get-Safe { $Item.Body } '')
        $flat = ''
        try { $flat = ([regex]::Replace($body, '\s+', ' ')).Trim() } catch { $flat = '' }
        $row['preview'] = Get-OlTruncated $flat $PreviewChars
    }

    return $row
}

# -------------------------------------------------------------------- doctor

function Get-OlStoreDefaultsMap {
    param($Store)
    $map = @{}
    $pairs = @(
        @('inbox', 6),
        @('drafts', 16),
        @('sent items', 5),
        @('deleted items', 3),
        @('junk email', 23),
        @('outbox', 4)
    )
    foreach ($pair in $pairs) {
        $label = [string]$pair[0]
        $code = [int]$pair[1]
        $folder = Get-OlStoreDefaultFolder $Store $code
        if ($null -eq $folder) { continue }
        $nm = $null
        try { $nm = [string]$folder.Name } catch { $nm = $null }
        if (-not [string]::IsNullOrEmpty($nm)) { $map[$label] = $nm }
    }
    # Archive is reported by name, not by type -- see the OL_FOLDER note.
    $root = $null
    try { $root = $Store.GetRootFolder() } catch { $root = $null }
    if ($null -ne $root) {
        foreach ($candidate in @('Archive', 'Archiv', 'Archiv1', ('Arch' + [char]0x00ED + 'v'))) {
            $found = Find-OlChild $root $candidate
            if ($null -ne $found) {
                $nm2 = $null
                try { $nm2 = [string]$found.Name } catch { $nm2 = $null }
                if (-not [string]::IsNullOrEmpty($nm2)) { $map['archive'] = $nm2 }
                break
            }
        }
    }
    return $map
}

function Invoke-OlDoctor {
    param($Ol, $Ns, $OpArgs)

    $quick = Get-OlBool $OpArgs 'quick' $false
    $stores = New-Object System.Collections.Generic.List[object]
    $n = 0
    try { $n = [int]$Ns.Folders.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $root = $null
        try { $root = $Ns.Folders.Item($i) } catch { $root = $null }
        if ($null -eq $root) { continue }
        $storeObj = Get-Safe { $root.Store } $null

        $entry = @{}
        $entry['name'] = Get-Safe { $root.Name } $null

        $storeId = $null
        if ($null -ne $storeObj) { $storeId = Get-Safe { $storeObj.StoreID } $null }
        if ($null -eq $storeId) { $storeId = Get-Safe { $root.StoreID } $null }
        $entry['store_id'] = $storeId

        $kind = $null
        if ($null -ne $storeObj) {
            $raw = Get-Safe { $storeObj.ExchangeStoreType } $null
            if ($null -ne $raw) {
                $asInt = $null
                try { $asInt = [int]$raw } catch { $asInt = $null }
                if (($null -ne $asInt) -and $script:EXCHANGE_STORE_TYPE.ContainsKey($asInt)) {
                    $kind = $script:EXCHANGE_STORE_TYPE[$asInt]
                } elseif ($null -ne $asInt) {
                    $kind = $asInt
                } else {
                    $kind = $raw
                }
            }
        }
        $entry['exchange_store_type'] = $kind

        # $null means Outlook would not answer, which is NOT the same as false.
        $cached = $null
        if ($null -ne $storeObj) {
            try { $cached = [bool]$storeObj.IsCachedExchange } catch { $cached = $null }
        }
        $entry['cached_mode'] = $cached

        # FilePath throws in online mode -- that absence is itself the signal.
        $dataFile = $null
        if ($null -ne $storeObj) {
            try {
                $fp = [string]$storeObj.FilePath
                if (-not [string]::IsNullOrEmpty($fp)) { $dataFile = [System.IO.Path]::GetFileName($fp) }
            } catch { $dataFile = $null }
        }
        $entry['data_file'] = $dataFile

        # --quick skips the per-store folder walk: on online-mode stores each
        # of those lookups is a server round trip.
        if (-not $quick) {
            $entry['default_folders'] = Get-OlStoreDefaultsMap $storeObj
            $rootChildren = Get-OlChildNames $root
            $entry['top_level_folders'] = @($rootChildren)
        }

        [void]$stores.Add($entry)
    }

    if ($quick) {
        $data = @{}
        $data['outlook_version'] = Get-Safe { $Ol.Version } $null
        $data['profile'] = Get-Safe { $Ns.CurrentProfileName } $null
        $data['current_user'] = Get-Safe { $Ns.CurrentUser.Name } $null
        $data['stores'] = @($stores.ToArray())
        $data['quick'] = $true
        return $data
    }

    $protected = Get-OlProtectedFolderIds $Ns
    $protValues = New-Object System.Collections.Generic.List[object]
    foreach ($v in (@($protected.Values) | Sort-Object)) { [void]$protValues.Add($v) }

    $data = @{}
    $data['outlook_version'] = Get-Safe { $Ol.Version } $null
    $data['profile'] = Get-Safe { $Ns.CurrentProfileName } $null
    $data['current_user'] = Get-Safe { $Ns.CurrentUser.Name } $null
    $data['default_inbox'] = Get-Safe { Get-OlFolderPath -Folder $Ns.GetDefaultFolder(6) } $null
    $data['stores'] = @($stores.ToArray())
    $data['protected_folders_resolved'] = @($protValues.ToArray())
    return $data
}

# ------------------------------------------------------------------- folders

function Add-OlFolderTree {
    param($Folder, [int]$Depth, [int]$MaxDepth, $Sink)
    if ($null -eq $Folder) { return }

    $count = $null
    try { $count = [int]$Folder.Items.Count } catch { $count = $null }

    $unread = $null
    try { $unread = [int]$Folder.UnReadItemCount } catch { $unread = $null }

    $entry = @{}
    $entry['name'] = Get-Safe { $Folder.Name } $null
    $entry['path'] = Get-OlFolderPath -Folder $Folder
    $entry['items'] = $count
    $entry['unread'] = $unread
    $entry['depth'] = $Depth
    [void]$Sink.Add($entry)

    if ($Depth -ge $MaxDepth) { return }

    $n = 0
    try { $n = [int]$Folder.Folders.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $sub = $null
        try { $sub = $Folder.Folders.Item($i) } catch { $sub = $null }
        if ($null -eq $sub) { continue }
        Add-OlFolderTree -Folder $sub -Depth ($Depth + 1) -MaxDepth $MaxDepth -Sink $Sink
    }
}

function Invoke-OlFolders {
    param($Ol, $Ns, $OpArgs)

    $folderArg = Get-OlString $OpArgs 'folder' $null
    $sharedArg = Get-OlString $OpArgs 'shared' $null
    $depth = Get-OlInt $OpArgs 'depth' 2

    $roots = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrEmpty($folderArg) -and [string]::IsNullOrEmpty($sharedArg)) {
        $n = 0
        try { $n = [int]$Ns.Folders.Count } catch { $n = 0 }
        for ($i = 1; $i -le $n; $i++) {
            $r = $null
            try { $r = $Ns.Folders.Item($i) } catch { $r = $null }
            if ($null -ne $r) { [void]$roots.Add($r) }
        }
    } else {
        $p = $folderArg
        if ($null -eq $p) { $p = '' }
        [void]$roots.Add((Resolve-OlFolder -Ns $Ns -Path $p -Shared $sharedArg))
    }

    $sink = New-Object System.Collections.Generic.List[object]
    foreach ($root in $roots) {
        Add-OlFolderTree -Folder $root -Depth 0 -MaxDepth $depth -Sink $sink
    }

    $data = @{}
    $data['count'] = $sink.Count
    $data['folders'] = @($sink.ToArray())
    return $data
}

# ---------------------------------------------------------------------- list

function Test-OlRowFilters {
    param($Row, $Filters)
    if ($null -eq $Filters) { return $true }

    $sender = Get-OlString $Filters 'sender' $null
    if (-not [string]::IsNullOrEmpty($sender)) {
        $email = [string]$Row['from_email']
        $name = [string]$Row['from_name']
        $hay = ($email + ' ' + $name).ToLower()
        if ($hay.IndexOf($sender.ToLower()) -lt 0) { return $false }
    }

    $subject = Get-OlString $Filters 'subject' $null
    if (-not [string]::IsNullOrEmpty($subject)) {
        $hay = ([string]$Row['subject']).ToLower()
        if ($hay.IndexOf($subject.ToLower()) -lt 0) { return $false }
    }

    if (Get-OlBool $Filters 'has_attachments' $false) {
        if ([int]$Row['attachment_count'] -le 0) { return $false }
    }

    $category = Get-OlString $Filters 'category' $null
    if (-not [string]::IsNullOrEmpty($category)) {
        $want = $category.ToLower()
        $hit = $false
        foreach ($c in @($Row['categories'])) {
            if (([string]$c).ToLower() -eq $want) { $hit = $true; break }
        }
        if (-not $hit) { return $false }
    }

    return $true
}

# Noise exclusions (--exclude-sender, --exclude-subject, --preset). Kept apart
# from Test-OlRowFilters so the listing can report how many rows they removed.
function Test-OlRowExcluded {
    param($Row, $Filters)
    if ($null -eq $Filters) { return $false }
    $hay = (([string]$Row['from_email']) + ' ' + ([string]$Row['from_name'])).ToLower()
    foreach ($s in (Get-OlArgArray $Filters 'exclude_sender')) {
        $w = ([string]$s).Trim().ToLower()
        if (($w.Length -gt 0) -and ($hay.IndexOf($w) -ge 0)) { return $true }
    }
    $subj = ([string]$Row['subject']).ToLower()
    foreach ($s in (Get-OlArgArray $Filters 'exclude_subject')) {
        $w = ([string]$s).Trim().ToLower()
        if (($w.Length -gt 0) -and ($subj.IndexOf($w) -ge 0)) { return $true }
    }
    return $false
}

# Mail folders under $Folder (not $Folder itself), depth-first, capped so a
# huge tree cannot turn one list into hundreds of server round trips.
function Add-OlMailSubfolders {
    param($Folder, $Sink, [int]$Cap = 100)
    $n = 0
    try { $n = [int]$Folder.Folders.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        if ($Sink.Count -ge $Cap) { return }
        $sub = $null
        try { $sub = $Folder.Folders.Item($i) } catch { $sub = $null }
        if ($null -eq $sub) { continue }
        $kind = 0
        try { $kind = [int]$sub.DefaultItemType } catch { $kind = 0 }
        if ($kind -ne 0) { continue }
        [void]$Sink.Add($sub)
        Add-OlMailSubfolders -Folder $sub -Sink $Sink -Cap $Cap
    }
}

<#
    One folder's share of a listing. Items are walked in received order in
    blocks of equal received-second, each block sorted by ref, so a cursor can
    sit between two mails with the same timestamp. Stops at the first item
    beyond --limit or --scan-max and says which one stopped it.
#>
function Get-OlFolderPage {
    param($Folder, [string]$Restrict, [int]$Limit, [int]$ScanMax, [bool]$OldestFirst,
        [int]$PreviewChars, [bool]$AnyClass, $Filters, $Cursor)

    $folderPath = Get-OlFolderPath -Folder $Folder
    $items = $null
    try {
        $items = $Folder.Items
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
            -Message ("could not open the items of '" + [string]$folderPath + "'")
    }

    $coll = $items
    if (-not [string]::IsNullOrEmpty($Restrict)) {
        try {
            $coll = $items.Restrict($Restrict)
        } catch {
            Invoke-OlFail -Code 'bad_filter' -Detail $_.Exception.Message `
                -Message ("Outlook rejected the filter '" + $Restrict + "'")
        }
    }

    try { $coll.Sort('[ReceivedTime]', (-not $OldestFirst)) } catch { }

    # Every item in a folder shares the folder's StoreID; read it once.
    $storeId = ''
    try { $storeId = [string]$Folder.StoreID } catch { $storeId = '' }

    $cmp = Get-OlPosComparison $OldestFirst
    $rows = New-Object System.Collections.Generic.List[object]
    $scanned = 0
    $skipped = 0
    $excluded = 0
    $truncBy = $null
    $frontier = $null
    $done = $false

    $it = $null
    try { $it = $coll.GetFirst() } catch { $it = $null }

    while (($null -ne $it) -and (-not $done)) {
        $block = New-Object System.Collections.Generic.List[object]
        $blockR = $null
        while ($null -ne $it) {
            $r = Get-OlIso (Get-Safe { $it.ReceivedTime } $null)
            if ($null -eq $r) { $r = '' }
            if (($block.Count -gt 0) -and ($r -ne $blockR)) { break }
            $blockR = $r
            $eid = ''
            try { $eid = [string]$it.EntryID } catch { $eid = '' }
            [void]$block.Add(@{ r = $r; ref = ($eid + '!' + $storeId); item = $it })
            try { $it = $coll.GetNext() } catch { $it = $null }
        }
        $block.Sort($cmp)

        foreach ($e in $block) {
            if (($null -ne $Cursor) -and (-not (Test-OlPosAfter $e $Cursor $OldestFirst))) { continue }
            if ($rows.Count -ge $Limit) { $truncBy = 'limit'; $done = $true; break }
            if ($scanned -ge $ScanMax) { $truncBy = 'scan_max'; $done = $true; break }
            $scanned = $scanned + 1
            $frontier = @{ r = $e['r']; ref = $e['ref'] }
            $cur = $e['item']
            try {
                $cls = 0
                try { $cls = [int](Get-Safe { $cur.Class } 0) } catch { $cls = 0 }
                if (($cls -ne $script:OL_MAIL_ITEM) -and (-not $AnyClass)) {
                    $skipped = $skipped + 1
                } else {
                    $row = Get-OlRow -Item $cur -PreviewChars $PreviewChars
                    if (Test-OlRowExcluded -Row $row -Filters $Filters) {
                        $excluded = $excluded + 1
                    } elseif (Test-OlRowFilters -Row $row -Filters $Filters) {
                        [void]$rows.Add($row)
                    }
                }
            } catch {
                $skipped = $skipped + 1
            }
        }
    }

    $page = @{}
    $page['folder'] = $folderPath
    $page['rows'] = $rows
    $page['scanned'] = $scanned
    $page['skipped'] = $skipped
    $page['excluded'] = $excluded
    $page['truncated_by'] = $truncBy
    $page['frontier'] = $frontier
    return $page
}

function Invoke-OlList {
    param($Ol, $Ns, $OpArgs)

    $folderList = Get-OlArgArray $OpArgs 'folders'
    $sharedArg = Get-OlString $OpArgs 'shared' $null
    $restrict = Get-OlString $OpArgs 'restrict' $null
    $limit = Get-OlInt $OpArgs 'limit' 25
    $scanMax = Get-OlInt $OpArgs 'scan_max' 2000
    $oldestFirst = Get-OlBool $OpArgs 'oldest_first' $false
    $previewChars = Get-OlInt $OpArgs 'preview_chars' 0
    $anyClass = Get-OlBool $OpArgs 'any_class' $false
    $countOnly = Get-OlBool $OpArgs 'count_only' $false
    $subfolders = Get-OlBool $OpArgs 'include_subfolders' $false
    $filters = Get-OlArg $OpArgs 'filters' $null
    $started = Get-OlString $OpArgs 'started' $null

    $cursor = $null
    $cursorArg = Get-OlArg $OpArgs 'cursor' $null
    if ($null -ne $cursorArg) {
        $cursor = @{ r = (Get-OlString $cursorArg 'r' ''); ref = (Get-OlString $cursorArg 'ref' '') }
    }

    $folders = New-Object System.Collections.Generic.List[object]
    foreach ($p in $folderList) {
        $f = Resolve-OlFolder -Ns $Ns -Path ([string]$p) -Shared $sharedArg
        [void]$folders.Add($f)
        if ($subfolders) { Add-OlMailSubfolders -Folder $f -Sink $folders }
    }

    $filterOut = $restrict
    if ([string]::IsNullOrEmpty($filterOut)) { $filterOut = $null }

    if ($countOnly) {
        $total = 0
        $per = New-Object System.Collections.Generic.List[object]
        foreach ($f in $folders) {
            $coll = $null
            try {
                $coll = $f.Items
                if (-not [string]::IsNullOrEmpty($restrict)) { $coll = $coll.Restrict($restrict) }
            } catch {
                Invoke-OlFail -Code 'bad_filter' -Detail $_.Exception.Message `
                    -Message ("Outlook rejected the filter '" + [string]$restrict + "'")
            }
            $c = 0
            try {
                $c = [int]$coll.Count
            } catch {
                Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not count items'
            }
            $total = $total + $c
            [void]$per.Add(@{ folder = (Get-OlFolderPath -Folder $f); count = $c })
        }
        $data = @{}
        $data['folder'] = Get-OlFolderPath -Folder $folders[0]
        $data['filter'] = $filterOut
        $data['count'] = $total
        if ($folders.Count -gt 1) { $data['per_folder'] = @($per.ToArray()) }
        return $data
    }

    $pages = New-Object System.Collections.Generic.List[object]
    foreach ($f in $folders) {
        $pg = Get-OlFolderPage -Folder $f -Restrict $restrict -Limit $limit -ScanMax $scanMax `
            -OldestFirst $oldestFirst -PreviewChars $previewChars -AnyClass $anyClass `
            -Filters $filters -Cursor $cursor
        [void]$pages.Add($pg)
    }

    # A folder that stopped early has not been read past its frontier. Nothing
    # beyond the earliest such frontier may be returned, or the next page
    # (which starts after the last returned row) would skip that folder's rest.
    $clip = $null
    $clipBy = $null
    foreach ($pg in $pages) {
        if (($null -ne $pg['truncated_by']) -and ($null -ne $pg['frontier'])) {
            if (($null -eq $clip) -or (Test-OlPosAfter $clip $pg['frontier'] $oldestFirst)) {
                $clip = $pg['frontier']
                $clipBy = $pg['truncated_by']
            }
        }
    }

    $merged = New-Object System.Collections.Generic.List[object]
    $scanned = 0
    $skipped = 0
    $excluded = 0
    foreach ($pg in $pages) {
        $scanned = $scanned + [int]$pg['scanned']
        $skipped = $skipped + [int]$pg['skipped']
        $excluded = $excluded + [int]$pg['excluded']
        foreach ($row in $pg['rows']) {
            $pos = @{ r = [string]$row['received']; ref = [string]$row['ref']; row = $row }
            if (($null -ne $clip) -and (Test-OlPosAfter $pos $clip $oldestFirst)) { continue }
            [void]$merged.Add($pos)
        }
    }
    $merged.Sort((Get-OlPosComparison $oldestFirst))

    $out = New-Object System.Collections.Generic.List[object]
    $truncBy = $null
    $next = $null
    for ($k = 0; $k -lt $merged.Count; $k++) {
        if ($out.Count -ge $limit) { $truncBy = 'limit'; break }
        [void]$out.Add($merged[$k]['row'])
        $next = $merged[$k]
    }
    if ($null -eq $truncBy) {
        if ($null -ne $clip) {
            $truncBy = $clipBy
            $next = $clip
        } else {
            $next = $null
        }
    }

    $data = @{}
    $data['folder'] = Get-OlFolderPath -Folder $folders[0]
    if ($folders.Count -gt 1) {
        $fl = New-Object System.Collections.Generic.List[object]
        foreach ($pg in $pages) { [void]$fl.Add($pg['folder']) }
        $data['folders'] = @($fl.ToArray())
    }
    $data['filter'] = $filterOut
    $data['returned'] = $out.Count
    $data['scanned'] = $scanned
    $data['skipped_non_mail_or_unreadable'] = $skipped
    $data['excluded'] = $excluded
    $data['truncated'] = ($null -ne $truncBy)
    $data['truncated_by'] = $truncBy
    $data['next_cursor'] = $null
    if ($null -ne $truncBy) {
        $data['next_cursor'] = ConvertTo-OlCursor -Pos $next -OldestFirst $oldestFirst -Started $started
    }
    $data['started'] = $started
    $data['items'] = @($out.ToArray())
    return $data
}

# ----------------------------------------------------------------------- get

function Invoke-OlGet {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $bodyMode = Get-OlString $OpArgs 'body' 'text'
    $bodyChars = Get-OlInt $OpArgs 'body_chars' 20000
    $wantHeaders = Get-OlBool $OpArgs 'headers' $false
    $stripQuoted = Get-OlBool $OpArgs 'strip_quoted' $false

    $item = Get-OlItem -Ns $Ns -Ref $ref
    $row = Get-OlRow -Item $item -PreviewChars 0

    $recipients = New-Object System.Collections.Generic.List[object]
    $rc = 0
    try { $rc = [int]$item.Recipients.Count } catch { $rc = 0 }
    for ($i = 1; $i -le $rc; $i++) {
        $rec = $null
        try { $rec = $item.Recipients.Item($i) } catch { $rec = $null }
        if ($null -eq $rec) { continue }
        $kind = 'to'
        $t = $null
        try { $t = [int](Get-Safe { $rec.Type } 1) } catch { $t = 1 }
        if (($null -ne $t) -and $script:RECIP_TYPE.ContainsKey($t)) { $kind = $script:RECIP_TYPE[$t] }
        $entry = @{}
        $entry['kind'] = $kind
        $entry['name'] = Get-Safe { $rec.Name } $null
        $entry['email'] = Get-OlRecipientSmtp -Recipient $rec
        [void]$recipients.Add($entry)
    }
    $row['recipients'] = @($recipients.ToArray())

    $attachments = New-Object System.Collections.Generic.List[object]
    $ac = 0
    try { $ac = [int]$item.Attachments.Count } catch { $ac = 0 }
    for ($i = 1; $i -le $ac; $i++) {
        $att = $null
        try { $att = $item.Attachments.Item($i) } catch { $att = $null }
        if ($null -eq $att) { continue }
        $nm = Get-Safe { $att.FileName } $null
        if ([string]::IsNullOrEmpty([string]$nm)) { $nm = Get-Safe { $att.DisplayName } $null }
        $entry = @{}
        $entry['index'] = $i
        $entry['name'] = $nm
        $sz = 0
        try { $sz = [int](Get-Safe { $att.Size } 0) } catch { $sz = 0 }
        $entry['size'] = $sz
        $ty = 0
        try { $ty = [int](Get-Safe { $att.Type } 0) } catch { $ty = 0 }
        $entry['type'] = $ty
        [void]$attachments.Add($entry)
    }
    $row['attachments'] = @($attachments.ToArray())

    if ($bodyMode -eq 'text') {
        $body = [string](Get-Safe { $item.Body } '')
        if ($stripQuoted) {
            $st = Get-OlStrippedBody $body
            $body = [string]$st['text']
            $row['quoted_chars_removed'] = $st['removed']
        }
        $row['body'] = Get-OlTruncated $body $bodyChars
        $row['body_truncated'] = ($body.Length -gt $bodyChars)
    } elseif ($bodyMode -eq 'html') {
        $body = [string](Get-Safe { $item.HTMLBody } '')
        $row['body_html'] = Get-OlTruncated $body $bodyChars
        $row['body_truncated'] = ($body.Length -gt $bodyChars)
    } else {
        $row['body_truncated'] = $false
    }

    if ($wantHeaders) {
        $row['internet_headers'] = Get-Safe { $item.PropertyAccessor.GetProperty($script:PR_INTERNET_HEADERS) } $null
    }

    $row['sent'] = Get-OlIso (Get-Safe { $item.SentOn } $null)
    $row['reply_to'] = Get-Safe { $item.ReplyRecipientNames } $null
    # Who a draft will go out as; an unsent item has no sender yet.
    $row['sent_on_behalf_of'] = Get-Safe { $item.SentOnBehalfOfName } $null
    $row['send_using_account'] = Get-Safe { [string]$item.SendUsingAccount.SmtpAddress } $null
    $row['internet_message_id'] = Get-Safe { $item.PropertyAccessor.GetProperty($script:PR_INTERNET_MESSAGE_ID) } $null

    return $row
}

# -------------------------------------------------------------------- thread

function Get-OlStoreFolderByCode {
    param($Item, [int]$Code)
    $store = Get-Safe { $Item.Parent.Store } $null
    return (Get-OlStoreDefaultFolder $store $Code)
}

<#
    Every message of one conversation, oldest first. Outlook's own
    conversation view (GetConversation) covers the item's store across
    folders, Sent Items included. A topic search over the item's folder,
    its store's Inbox and Sent Items and any --also folders adds the rest:
    another store's Sent Items is where replies sent from a personal
    mailbox end up. Topic hits only count when their ConversationID
    matches, so an unrelated mail with the same subject stays out.
#>
function Invoke-OlThread {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $bodyMode = Get-OlString $OpArgs 'body' 'text'
    $bodyChars = Get-OlInt $OpArgs 'body_chars' 4000
    $stripQuoted = Get-OlBool $OpArgs 'strip_quoted' $true
    $also = Get-OlArgArray $OpArgs 'also'
    $restrictExtra = Get-OlString $OpArgs 'restrict' $null

    $item = Get-OlItem -Ns $Ns -Ref $ref
    $convId = [string](Get-Safe { $item.ConversationID } '')
    $topic = [string](Get-Safe { $item.ConversationTopic } '')

    $found = @{}
    $topicOnly = @{}
    $methods = New-Object System.Collections.Generic.List[object]
    $found[(Get-OlRef -Item $item)] = $item

    $conv = $null
    try { $conv = $item.GetConversation() } catch { $conv = $null }
    if ($null -ne $conv) {
        $storeId = Get-Safe { [string]$item.Parent.StoreID } ''
        try {
            $tbl = $conv.GetTable()
            while (-not $tbl.EndOfTable) {
                $tr = $tbl.GetNextRow()
                $eid = $null
                try { $eid = [string]$tr.Item('EntryID') } catch { $eid = $null }
                if ([string]::IsNullOrEmpty($eid)) { continue }
                $x = $null
                try { $x = $Ns.GetItemFromID($eid, $storeId) } catch { $x = $null }
                if ($null -eq $x) { continue }
                $found[(Get-OlRef -Item $x)] = $x
            }
            [void]$methods.Add('conversation')
        } catch { }
    }

    $searchIn = New-Object System.Collections.Generic.List[object]
    foreach ($f in @((Get-Safe { $item.Parent } $null), (Get-OlStoreFolderByCode $item 6), (Get-OlStoreFolderByCode $item 5))) {
        if ($null -ne $f) { [void]$searchIn.Add($f) }
    }
    foreach ($p in $also) {
        [void]$searchIn.Add((Resolve-OlFolder -Ns $Ns -Path ([string]$p) -Shared $null))
    }

    # A client that restarts the thread on reply makes "AW: <subject>" the
    # topic, so the bare subject is searched as well.
    $topics = New-Object System.Collections.Generic.List[object]
    foreach ($cand in @($topic, (Get-Safe { [string]$item.Subject } ''))) {
        foreach ($t in @($cand, ([regex]::Replace([string]$cand, '^(\s*(re|aw|fw|fwd|wg|odp|vs|sv|tr)\s*:\s*)+', '', 'IgnoreCase')))) {
            $t = ([string]$t).Trim()
            if (($t.Length -gt 0) -and (-not $topics.Contains($t))) { [void]$topics.Add($t) }
        }
    }

    foreach ($tp in $topics) {
        $flt = "[ConversationTopic] = '" + $tp.Replace("'", "''") + "'"
        if (-not [string]::IsNullOrEmpty($restrictExtra)) { $flt = $flt + ' AND ' + $restrictExtra }
        $seen = @{}
        foreach ($f in $searchIn) {
            $fid = Get-Safe { [string]$f.EntryID } ''
            if ($seen.ContainsKey($fid)) { continue }
            $seen[$fid] = $true
            $coll = $null
            try { $coll = $f.Items.Restrict($flt) } catch { $coll = $null }
            if ($null -eq $coll) { continue }
            $x = $null
            try { $x = $coll.GetFirst() } catch { $x = $null }
            $guard = 0
            while (($null -ne $x) -and ($guard -lt 500)) {
                $guard = $guard + 1
                $xr = Get-OlRef -Item $x
                if (-not $found.ContainsKey($xr)) {
                    $found[$xr] = $x
                    # Some clients start a fresh thread index on reply, so a
                    # topic match with another ConversationID is kept but
                    # labelled: the reader decides whether it belongs.
                    $xc = [string](Get-Safe { $x.ConversationID } '')
                    if ([string]::IsNullOrEmpty($xc) -or ($xc -ne $convId)) { $topicOnly[$xr] = $true }
                }
                try { $x = $coll.GetNext() } catch { $x = $null }
            }
        }
    }
    if ($topics.Count -gt 0) { [void]$methods.Add('topic') }

    $msgs = New-Object System.Collections.Generic.List[object]
    foreach ($k in @($found.Keys)) {
        $x = $found[$k]
        $cls = 0
        try { $cls = [int](Get-Safe { $x.Class } 0) } catch { $cls = 0 }
        if ($cls -ne $script:OL_MAIL_ITEM) { continue }
        # Drafts sit in the conversation too; they are not part of what was said.
        $wasSent = $true
        try { $wasSent = [bool]$x.Sent } catch { $wasSent = $true }
        if (-not $wasSent) { continue }
        $row = Get-OlRow -Item $x -PreviewChars 0
        $row['sent'] = Get-OlIso (Get-Safe { $x.SentOn } $null)
        $row['match'] = 'conversation'
        if ($topicOnly.ContainsKey($k)) { $row['match'] = 'topic' }
        if ($bodyMode -eq 'text') {
            $body = [string](Get-Safe { $x.Body } '')
            if ($stripQuoted) {
                $st = Get-OlStrippedBody $body
                $body = [string]$st['text']
                $row['quoted_chars_removed'] = $st['removed']
            }
            $row['body'] = Get-OlTruncated $body $bodyChars
            $row['body_truncated'] = ($body.Length -gt $bodyChars)
        }
        [void]$msgs.Add(@{ r = [string]$row['sent']; ref = [string]$row['ref']; row = $row })
    }
    $msgs.Sort((Get-OlPosComparison $true))

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $msgs) { [void]$rows.Add($m['row']) }

    $data = @{}
    $data['ref'] = $ref
    $data['conversation_topic'] = $topic
    $data['conversation_id'] = $convId
    $data['method'] = @($methods.ToArray())
    $data['count'] = $rows.Count
    $data['messages'] = @($rows.ToArray())
    return $data
}

# --------------------------------------------------------- save_attachments

function Get-OlSafeFileName {
    param($Name, [string]$Fallback = 'attachment.bin')
    $n = ''
    if ($null -ne $Name) { $n = ([string]$Name).Trim() }
    if ($n.Length -eq 0) { $n = $Fallback }
    try { $n = [regex]::Replace($n, '[\x00-\x1f<>:"/\\|?*]', '_') } catch { }
    $n = $n.TrimEnd([char[]]@(' ', '.'))
    if ($n.Length -eq 0) { $n = $Fallback }
    if ($n.Length -gt 180) { $n = $n.Substring(0, 180) }
    return $n
}

function Invoke-OlSaveAttachments {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $out = Get-OlString $OpArgs 'out' $null
    $pattern = Get-OlString $OpArgs 'pattern' $null
    $overwrite = Get-OlBool $OpArgs 'overwrite' $false
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    if ([string]::IsNullOrEmpty($out)) {
        Invoke-OlFail -Code 'bad_argument' -Message 'an output directory is required (a Windows path, e.g. C:\Users\me\Downloads)'
    }

    $item = Get-OlItem -Ns $Ns -Ref $ref

    # The caller already translated this to a Windows path; use it as-is.
    $outDir = $out
    try { $outDir = [System.IO.Path]::GetFullPath($out) } catch { $outDir = $out }

    $total = 0
    try { $total = [int](Get-Safe { $item.Attachments.Count } 0) } catch { $total = 0 }

    if (-not $dryRun) {
        try {
            if (-not (Test-Path -LiteralPath $outDir)) {
                [void](New-Item -ItemType Directory -Path $outDir -Force)
            }
        } catch {
            Invoke-OlFail -Code 'no_access' -Detail $_.Exception.Message `
                -Message ("could not create the output directory '" + $outDir + "'")
        }
    }

    $saved = New-Object System.Collections.Generic.List[object]
    $skipped = New-Object System.Collections.Generic.List[object]

    for ($i = 1; $i -le $total; $i++) {
        $att = $null
        try { $att = $item.Attachments.Item($i) } catch { $att = $null }
        if ($null -eq $att) {
            $s = @{}
            $s['index'] = $i
            $s['name'] = $null
            $s['reason'] = 'unreadable'
            [void]$skipped.Add($s)
            continue
        }

        $raw = Get-Safe { $att.FileName } $null
        if ([string]::IsNullOrEmpty([string]$raw)) { $raw = Get-Safe { $att.DisplayName } $null }
        if ([string]::IsNullOrEmpty([string]$raw)) { $raw = ('part' + [string]$i) }
        $name = Get-OlSafeFileName $raw

        $attType = 1
        try { $attType = [int](Get-Safe { $att.Type } 1) } catch { $attType = 1 }
        if ($attType -eq 6) {
            # olOLE -- not a real file, SaveAsFile would produce nothing useful.
            $s = @{}
            $s['index'] = $i
            $s['name'] = $name
            $s['reason'] = 'ole_object'
            [void]$skipped.Add($s)
            continue
        }

        if (-not [string]::IsNullOrEmpty($pattern)) {
            if (-not ($name.ToLower() -like $pattern.ToLower())) {
                $s = @{}
                $s['index'] = $i
                $s['name'] = $name
                $s['reason'] = 'pattern'
                [void]$skipped.Add($s)
                continue
            }
        }

        $target = [System.IO.Path]::Combine($outDir, $name)
        if (-not $overwrite) {
            $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
            $ext = [System.IO.Path]::GetExtension($name)
            $k = 1
            while ($true) {
                $exists = $false
                try { $exists = [bool](Test-Path -LiteralPath $target) } catch { $exists = $false }
                if (-not $exists) { break }
                $target = [System.IO.Path]::Combine($outDir, ($stem + ' (' + [string]$k + ')' + $ext))
                $k = $k + 1
                if ($k -gt 10000) { break }
            }
        }

        if ($dryRun) {
            $e = @{}
            $e['index'] = $i
            $e['name'] = $name
            $e['path'] = $target
            $e['bytes'] = $null
            $e['dry_run'] = $true
            [void]$saved.Add($e)
            continue
        }

        try {
            $att.SaveAsFile($target)
            $bytes = $null
            try { $bytes = [int64](Get-Item -LiteralPath $target).Length } catch { $bytes = $null }
            $e = @{}
            $e['index'] = $i
            $e['name'] = $name
            $e['path'] = $target
            $e['bytes'] = $bytes
            $e['dry_run'] = $false
            [void]$saved.Add($e)
        } catch {
            $s = @{}
            $s['index'] = $i
            $s['name'] = $name
            $s['reason'] = [string]$_.Exception.Message
            [void]$skipped.Add($s)
        }
    }

    $data = @{}
    $data['ref'] = $ref
    $data['subject'] = Get-Safe { $item.Subject } $null
    $data['attachments'] = $total
    $data['saved'] = @($saved.ToArray())
    $data['skipped'] = @($skipped.ToArray())
    $data['out_dir'] = $outDir
    $data['dry_run'] = [bool]$dryRun
    return $data
}

# ---------------------------------------------------------------------- move

function Invoke-OlMove {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $to = Get-OlString $OpArgs 'to' $null
    $sharedArg = Get-OlString $OpArgs 'shared' $null
    $allowProtected = Get-OlBool $OpArgs 'allow_protected' $false
    $protectDefaults = Get-OlBool $OpArgs 'protect_defaults' $true
    $protectedNames = Get-OlArgArray $OpArgs 'protected_names'
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    if ([string]::IsNullOrEmpty($to)) {
        Invoke-OlFail -Code 'bad_argument' -Message 'a destination folder is required'
    }

    $item = Get-OlItem -Ns $Ns -Ref $ref
    $dest = Resolve-OlFolder -Ns $Ns -Path $to -Shared $sharedArg
    $destName = [string](Get-Safe { $dest.Name } '?')
    $destPath = Get-OlFolderPath -Folder $dest

    if (-not $allowProtected) {
        $reason = $null
        if ($protectDefaults) {
            $entry = $null
            try { $entry = [string]$dest.EntryID } catch { $entry = $null }
            if (-not [string]::IsNullOrEmpty($entry)) {
                $ids = Get-OlProtectedFolderIds $Ns
                if ($ids.ContainsKey($entry)) {
                    $reason = ('it is a protected default folder (' + [string]$ids[$entry] + ')')
                }
            }
        }
        if ($null -eq $reason) {
            $lower = $destName.Trim().ToLower()
            foreach ($p in $protectedNames) {
                if (([string]$p).Trim().ToLower() -eq $lower) {
                    $reason = 'its name is in the configured protected_folders list'
                    break
                }
            }
        }
        if ($null -ne $reason) {
            Invoke-OlFail -Code 'protected_folder' -Message ("destination folder '" + $destName + "' is protected: " + $reason +
                '. Pass --allow-protected if that is really intended.')
        }
    }

    $src = Get-Safe { Get-OlFolderPath -Folder $item.Parent } $null
    $subject = Get-Safe { $item.Subject } $null

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['ref'] = $ref
        $data['subject'] = $subject
        $data['from'] = $src
        $data['to'] = $destPath
        return $data
    }

    $moved = $null
    try {
        $moved = $item.Move($dest)
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
            -Message ("Outlook refused to move the item to '" + [string]$destPath + "'")
    }

    # The EntryID legitimately changes on a move; the old ref is now stale.
    $newRef = Get-Safe { Get-OlRef -Item $moved } $null

    $data = @{}
    $data['subject'] = $subject
    $data['from'] = $src
    $data['to'] = $destPath
    $data['old_ref'] = $ref
    $data['new_ref'] = $newRef
    $data['note'] = 'the ref changed; use new_ref for further commands'
    $data['dry_run'] = $false
    return $data
}

# ---------------------------------------------------------------------- mark

function Invoke-OlMark {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $unread = Get-OlBool $OpArgs 'unread' $false
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    $item = Get-OlItem -Ns $Ns -Ref $ref
    $subject = Get-Safe { $item.Subject } $null

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['ref'] = $ref
        $data['subject'] = $subject
        $data['unread'] = [bool]$unread
        return $data
    }

    try {
        $item.UnRead = $unread
        $item.Save()
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not change the read state of the item'
    }

    $data = @{}
    $data['subject'] = $subject
    $data['unread'] = [bool](Get-Safe { $item.UnRead } $unread)
    $data['ref'] = $ref
    $data['dry_run'] = $false
    return $data
}

# ---------------------------------------------------------------------- flag

function Invoke-OlFlag {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $clear = Get-OlBool $OpArgs 'clear' $false
    $text = Get-OlString $OpArgs 'text' 'Follow up'
    $dueCode = Get-OlInt $OpArgs 'due_code' 3
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    $item = Get-OlItem -Ns $Ns -Ref $ref
    $subject = Get-Safe { $item.Subject } $null

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['ref'] = $ref
        $data['subject'] = $subject
        $data['flag'] = Get-OlFlagState -Item $item
        $data['clear'] = [bool]$clear
        $data['text'] = $text
        return $data
    }

    try {
        if ($clear) {
            try {
                $item.ClearTaskFlag()
            } catch {
                $item.FlagStatus = 0
            }
        } else {
            try { $item.MarkAsTask($dueCode) } catch { }
            $item.FlagRequest = $text
            $item.FlagStatus = 2
        }
        $item.Save()
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not change the follow-up flag on the item'
    }

    $data = @{}
    $data['subject'] = $subject
    $data['flag'] = Get-OlFlagState -Item $item
    $data['ref'] = $ref
    $data['dry_run'] = $false
    return $data
}

# ---------------------------------------------------------------- categorize

function Invoke-OlCategorize {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false
    $hasSet = Test-OlArgPresent $OpArgs 'set'
    $setList = Get-OlArgArray $OpArgs 'set'
    $addList = Get-OlArgArray $OpArgs 'add'
    $removeList = Get-OlArgArray $OpArgs 'remove'

    $item = Get-OlItem -Ns $Ns -Ref $ref
    $subject = Get-Safe { $item.Subject } $null

    $before = New-Object System.Collections.Generic.List[object]
    foreach ($c in (Get-OlCategoryList -Item $item)) { [void]$before.Add([string]$c) }

    $after = New-Object System.Collections.Generic.List[object]
    if ($hasSet) {
        foreach ($c in $setList) { [void]$after.Add([string]$c) }
    } else {
        foreach ($c in $before) { [void]$after.Add([string]$c) }
        foreach ($a in $addList) {
            $cand = [string]$a
            if ([string]::IsNullOrEmpty($cand)) { continue }
            $seen = $false
            foreach ($existing in $after) {
                if (([string]$existing).ToLower() -eq $cand.ToLower()) { $seen = $true; break }
            }
            if (-not $seen) { [void]$after.Add($cand) }
        }
        foreach ($r in $removeList) {
            $victim = ([string]$r).ToLower()
            $kept = New-Object System.Collections.Generic.List[object]
            foreach ($existing in $after) {
                if (([string]$existing).ToLower() -ne $victim) { [void]$kept.Add($existing) }
            }
            $after = $kept
        }
    }

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['ref'] = $ref
        $data['subject'] = $subject
        $data['before'] = @($before.ToArray())
        $data['after'] = @($after.ToArray())
        return $data
    }

    try {
        $item.Categories = ([string]::Join(';', @($after.ToArray())))
        $item.Save()
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not write the categories back to the item'
    }

    $data = @{}
    $data['subject'] = $subject
    $data['before'] = @($before.ToArray())
    $data['after'] = @($after.ToArray())
    $data['ref'] = $ref
    $data['dry_run'] = $false
    return $data
}

# ------------------------------------------------------------ draft helpers

# Who a draft goes out as. A profile account with that address is used
# directly; anything else (a shared mailbox mounted as a store) becomes
# Send-on-behalf, which Exchange turns into Send-as when the right exists.
function Set-OlDraftSender {
    param($Ns, $Mail, [string]$From)
    if ([string]::IsNullOrEmpty($From)) { return $null }
    $want = $From.Trim().ToLower()
    $n = 0
    try { $n = [int]$Ns.Accounts.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $acct = $null
        try { $acct = $Ns.Accounts.Item($i) } catch { $acct = $null }
        if ($null -eq $acct) { continue }
        $addr = ''
        try { $addr = ([string]$acct.SmtpAddress).Trim().ToLower() } catch { $addr = '' }
        if ($addr -eq $want) {
            try {
                $Mail.SendUsingAccount = $acct
                return 'account'
            } catch { }
        }
    }
    try {
        $Mail.SentOnBehalfOfName = $From
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message ("could not set the sender to '" + $From + "'")
    }
    return 'on_behalf'
}

# Store display name of the store an item lives in.
function Get-OlItemStoreName {
    param($Item)
    $nm = Get-Safe { [string]$Item.Parent.Store.DisplayName } $null
    if ([string]::IsNullOrEmpty($nm)) { $nm = Get-Safe { [string]$Item.Parent.Store.GetRootFolder().Name } $null }
    return $nm
}

# from_by_store from config.json: {"<store name>": "<address>"}.
function Get-OlFromForStore {
    param($Map, [string]$StoreName)
    if (($null -eq $Map) -or [string]::IsNullOrEmpty($StoreName)) { return $null }
    foreach ($p in $Map.PSObject.Properties) {
        if (([string]$p.Name).Trim().ToLower() -eq $StoreName.Trim().ToLower()) { return [string]$p.Value }
    }
    return $null
}

# Where the draft is going and from whom, worked out once so --dry-run reports
# exactly what a real run would do. An explicit --from / --save-to wins; else a
# from_by_store entry for the source store supplies both.
function Get-OlDraftPlan {
    param($Ns, $OpArgs, $SourceItem)
    $from = Get-OlString $OpArgs 'from' $null
    $saveTo = Get-OlString $OpArgs 'save_to' $null
    if ($null -ne $SourceItem) {
        $storeName = Get-OlItemStoreName $SourceItem
        $mapped = Get-OlFromForStore (Get-OlArg $OpArgs 'from_by_store' $null) $storeName
        if (-not [string]::IsNullOrEmpty($mapped)) {
            if ([string]::IsNullOrEmpty($from)) { $from = $mapped }
            if ([string]::IsNullOrEmpty($saveTo)) { $saveTo = ($storeName.Replace('/', '\/') + '/drafts') }
        }
    }
    $folder = $null
    if (-not [string]::IsNullOrEmpty($saveTo)) {
        $folder = Resolve-OlFolder -Ns $Ns -Path $saveTo -Shared $null
    }
    $plan = @{}
    $plan['from'] = $from
    $plan['folder'] = $folder
    $plan['folder_path'] = $null
    if ($null -ne $folder) { $plan['folder_path'] = Get-OlFolderPath -Folder $folder }
    return $plan
}

# After Save() a new item sits in the default store's Drafts. Move it where the
# plan says; the moved copy is the draft from now on (its ref changes).
function Move-OlDraftToPlan {
    param($Ns, $Mail, $Plan)
    if ($null -eq $Plan['folder']) { return $Mail }
    $here = Get-Safe { [string]$Mail.Parent.EntryID } ''
    $there = Get-Safe { [string]$Plan['folder'].EntryID } ''
    if ((-not [string]::IsNullOrEmpty($here)) -and ($here -eq $there)) { return $Mail }
    try {
        return $Mail.Move($Plan['folder'])
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
            -Message ("the draft was saved, but could not be moved to '" + [string]$Plan['folder_path'] + "'; it is in the default Drafts")
    }
}

# Where $Prefix ends inside $Text, ignoring whitespace differences, or -1 if
# $Text does not start with it. Outlook rewrites whitespace on save (drops a
# trailing newline, re-wraps an HTML reply's line breaks), so an exact
# StartsWith never matches the text it was given.
function Get-OlTextPrefixEnd {
    param([string]$Text, [string]$Prefix)
    $i = 0
    $j = 0
    $end = 0
    while ($j -lt $Prefix.Length) {
        if ([char]::IsWhiteSpace($Prefix[$j])) { $j++; continue }
        while (($i -lt $Text.Length) -and [char]::IsWhiteSpace($Text[$i])) { $i++ }
        if (($i -ge $Text.Length) -or ($Text[$i] -cne $Prefix[$j])) { return -1 }
        $i++
        $j++
        $end = $i
    }
    return $end
}

function Set-OlDraftTextMark {
    param($Mail, [string]$Text)
    try { $Mail.PropertyAccessor.SetProperty($script:OLCTL_DRAFT_TEXT, $Text) } catch { }
}

# An unsent mail item sitting in its own store's Drafts folder, or a failure.
function Assert-OlDraft {
    param($Item, [string]$Ref)
    $cls = 0
    try { $cls = [int]$Item.Class } catch { $cls = 0 }
    $sent = $true
    try { $sent = [bool]$Item.Sent } catch { $sent = $true }
    $inDrafts = $false
    $drafts = Get-OlStoreFolderByCode $Item 16
    if ($null -ne $drafts) {
        $inDrafts = ((Get-Safe { [string]$Item.Parent.EntryID } 'x') -eq (Get-Safe { [string]$drafts.EntryID } 'y'))
    }
    if (($cls -ne $script:OL_MAIL_ITEM) -or $sent -or (-not $inDrafts)) {
        $extra = @{}
        $extra['folder'] = Get-Safe { Get-OlFolderPath -Folder $Item.Parent } $null
        Invoke-OlFail -Code 'not_a_draft' -Extra $extra -Message 'that ref is not an unsent mail in a Drafts folder; only drafts can be updated or superseded'
    }
}

# --------------------------------------------------------------- draft_reply

function Invoke-OlDraftReply {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $text = Get-OlString $OpArgs 'text' ''
    if ($null -eq $text) { $text = '' }
    $replyAll = Get-OlBool $OpArgs 'reply_all' $false
    $asHtml = Get-OlBool $OpArgs 'html' $false
    $subjectOverride = Get-OlString $OpArgs 'subject' $null
    $attach = Get-OlArgArray $OpArgs 'attach'
    $supersede = Get-OlString $OpArgs 'supersede' $null
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    $item = Get-OlItem -Ns $Ns -Ref $ref
    $plan = Get-OlDraftPlan -Ns $Ns -OpArgs $OpArgs -SourceItem $item

    # Checked before anything is created, so a bad --supersede ref leaves no
    # half-done state behind.
    $old = $null
    $oldBin = $null
    if (-not [string]::IsNullOrEmpty($supersede)) {
        $old = Get-OlItem -Ns $Ns -Ref $supersede
        Assert-OlDraft -Item $old -Ref $supersede
        $oldBin = Get-OlStoreFolderByCode $old 3
        if ($null -eq $oldBin) {
            Invoke-OlFail -Code 'error' -Message 'the superseded draft has no Deleted Items folder to go to'
        }
    }

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['ref'] = $ref
        $data['reply_all'] = [bool]$replyAll
        $data['chars'] = $text.Length
        $data['subject'] = Get-Safe { $item.Subject } $null
        $data['attach'] = @($attach)
        $data['from'] = $plan['from']
        $data['saved_to'] = $plan['folder_path']
        if ($null -eq $data['saved_to']) { $data['saved_to'] = 'Drafts' }
        $data['supersedes'] = $supersede
        $data['sent'] = $false
        return $data
    }

    $reply = $null
    try {
        if ($replyAll) { $reply = $item.ReplyAll() } else { $reply = $item.Reply() }
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'Outlook could not build a reply for that item'
    }
    if ($null -eq $reply) {
        Invoke-OlFail -Code 'error' -Message 'Outlook could not build a reply for that item'
    }

    $sender = $null
    try {
        $sender = Set-OlDraftSender -Ns $Ns -Mail $reply -From $plan['from']
        if ($asHtml) {
            $existing = [string](Get-Safe { $reply.HTMLBody } '')
            $reply.HTMLBody = $text + $existing
        } else {
            # Plain text in an HTML item gets every <mailto:...> auto-linked,
            # and each read-back adds another copy; a plain-format draft
            # round-trips unchanged.
            $reply.BodyFormat = $script:OL_FORMAT_PLAIN
            $existing = [string](Get-Safe { $reply.Body } '')
            $reply.Body = $text + "`n" + $existing
        }
        if (-not [string]::IsNullOrEmpty($subjectOverride)) { $reply.Subject = $subjectOverride }
    } catch {
        if ($null -ne $script:OlErrorPayload) { throw }
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not save the reply draft'
    }

    foreach ($p in $attach) {
        $winPath = [string]$p
        if ([string]::IsNullOrEmpty($winPath)) { continue }
        try {
            [void]$reply.Attachments.Add($winPath)
        } catch {
            Invoke-OlFail -Code 'bad_argument' -Detail $_.Exception.Message `
                -Message ("could not attach '" + $winPath + "'. It must be a Windows path reachable from Outlook.")
        }
    }

    Set-OlDraftTextMark -Mail $reply -Text $text

    try {
        # Lands in Drafts. This tool never sends: there is no Send() call here.
        $reply.Save()
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not save the reply draft'
    }
    $reply = Move-OlDraftToPlan -Ns $Ns -Mail $reply -Plan $plan

    $superseded = $null
    if ($null -ne $old) {
        try {
            # Into Deleted Items, never a hard delete: recoverable from Outlook.
            [void]$old.Move($oldBin)
            $superseded = $supersede
        } catch {
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
                -Message ('the new draft was saved, but the old one could not be moved to Deleted Items; ' +
                    'new draft_ref: ' + [string](Get-Safe { Get-OlRef -Item $reply } '?'))
        }
    }

    $data = @{}
    $data['draft_ref'] = Get-Safe { Get-OlRef -Item $reply } $null
    $data['in_reply_to'] = $ref
    $data['subject'] = Get-Safe { $reply.Subject } $null
    $data['from'] = $plan['from']
    $data['from_mode'] = $sender
    $data['saved_to'] = Get-Safe { Get-OlFolderPath -Folder $reply.Parent } 'Drafts'
    $data['superseded'] = $superseded
    $data['sent'] = $false
    $data['dry_run'] = $false
    return $data
}

# ----------------------------------------------------------------- draft_new

function Invoke-OlDraftNew {
    param($Ol, $Ns, $OpArgs)

    $to = Get-OlArgArray $OpArgs 'to'
    $cc = Get-OlArgArray $OpArgs 'cc'
    $subject = Get-OlString $OpArgs 'subject' ''
    if ($null -eq $subject) { $subject = '' }
    $text = Get-OlString $OpArgs 'text' ''
    if ($null -eq $text) { $text = '' }
    $asHtml = Get-OlBool $OpArgs 'html' $false
    $attach = Get-OlArgArray $OpArgs 'attach'
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    $plan = Get-OlDraftPlan -Ns $Ns -OpArgs $OpArgs -SourceItem $null

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['to'] = @($to)
        $data['cc'] = @($cc)
        $data['subject'] = $subject
        $data['chars'] = $text.Length
        $data['attach'] = @($attach)
        $data['from'] = $plan['from']
        $data['saved_to'] = $plan['folder_path']
        if ($null -eq $data['saved_to']) { $data['saved_to'] = 'Drafts' }
        $data['sent'] = $false
        return $data
    }

    $mail = $null
    try {
        $mail = $Ol.CreateItem(0)
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not create a new mail item'
    }

    $sender = $null
    try {
        $sender = Set-OlDraftSender -Ns $Ns -Mail $mail -From $plan['from']
        $mail.To = [string]::Join('; ', @($to))
        if ($cc.Count -gt 0) { $mail.CC = [string]::Join('; ', @($cc)) }
        $mail.Subject = $subject
        if ($asHtml) {
            $mail.HTMLBody = $text
        } else {
            $mail.BodyFormat = $script:OL_FORMAT_PLAIN
            $mail.Body = $text
        }
    } catch {
        if ($null -ne $script:OlErrorPayload) { throw }
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not populate the new draft'
    }

    foreach ($p in $attach) {
        $winPath = [string]$p
        if ([string]::IsNullOrEmpty($winPath)) { continue }
        try {
            [void]$mail.Attachments.Add($winPath)
        } catch {
            Invoke-OlFail -Code 'bad_argument' -Detail $_.Exception.Message `
                -Message ("could not attach '" + $winPath + "'. It must be a Windows path reachable from Outlook.")
        }
    }

    Set-OlDraftTextMark -Mail $mail -Text $text

    try {
        # Lands in Drafts. This tool never sends: there is no Send() call here.
        $mail.Save()
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not save the new draft'
    }
    $mail = Move-OlDraftToPlan -Ns $Ns -Mail $mail -Plan $plan

    $data = @{}
    $data['draft_ref'] = Get-Safe { Get-OlRef -Item $mail } $null
    $data['subject'] = Get-Safe { $mail.Subject } $subject
    $data['from'] = $plan['from']
    $data['from_mode'] = $sender
    $data['saved_to'] = Get-Safe { Get-OlFolderPath -Folder $mail.Parent } 'Drafts'
    $data['sent'] = $false
    $data['dry_run'] = $false
    return $data
}

# -------------------------------------------------------------- draft_update

<#
    Rewrite an existing draft in place, so iterating on a reply does not leave
    a trail of near-identical drafts. By default only the text olctl wrote is
    swapped and the quoted original below it is kept; that needs the draft to
    carry the olctl_text mark and still start with that text. --whole-body
    replaces everything instead.
#>
function Invoke-OlDraftUpdate {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $hasText = Test-OlArgPresent $OpArgs 'text'
    $text = Get-OlString $OpArgs 'text' ''
    if ($null -eq $text) { $text = '' }
    $asHtml = Get-OlBool $OpArgs 'html' $false
    $wholeBody = Get-OlBool $OpArgs 'whole_body' $false
    $subject = Get-OlString $OpArgs 'subject' $null
    $from = Get-OlString $OpArgs 'from' $null
    $attach = Get-OlArgArray $OpArgs 'attach'
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    $item = Get-OlItem -Ns $Ns -Ref $ref
    Assert-OlDraft -Item $item -Ref $ref

    $newBody = $null
    $mode = 'unchanged'
    if ($hasText) {
        $prop = $null
        try { $prop = [string]$item.PropertyAccessor.GetProperty($script:OLCTL_DRAFT_TEXT) } catch { $prop = $null }
        $current = ''
        if ($asHtml) { $current = [string](Get-Safe { $item.HTMLBody } '') } else { $current = [string](Get-Safe { $item.Body } '') }
        # Outlook stores CRLF whatever was written; compare on LF.
        $curN = $current.Replace("`r`n", "`n")
        $oldN = ''
        if ($null -ne $prop) { $oldN = $prop.Replace("`r`n", "`n") }
        if ($wholeBody) {
            $newBody = $text
            $mode = 'whole_body'
        } elseif ((-not [string]::IsNullOrEmpty($oldN)) -and ((Get-OlTextPrefixEnd $curN $oldN) -ge 0)) {
            $newBody = $text + $curN.Substring((Get-OlTextPrefixEnd $curN $oldN))
            $mode = 'text_only'
        } else {
            $why = 'no olctl text mark on the draft'
            if (-not [string]::IsNullOrEmpty($oldN)) {
                $show = $curN
                if ($show.Length -gt 80) { $show = $show.Substring(0, 80) }
                $why = ('marked text (' + [string]$oldN.Length + ' chars) is not at the start of the body, which begins: ' + $show)
            }
            Invoke-OlFail -Code 'cannot_locate_text' -Detail $why -Message ('cannot tell which part of this draft is the reply text: ' +
                'it was not written by olctl 0.4+, or it was edited in Outlook since. Pass --whole-body to replace ' +
                'the entire body (the quoted original is lost), or rebuild it with draft-reply --supersede <draft_ref>.')
        }
    }

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['draft_ref'] = $ref
        $data['subject'] = Get-Safe { $item.Subject } $null
        $data['new_subject'] = $subject
        $data['body'] = $mode
        $data['chars'] = $text.Length
        $data['from'] = $from
        $data['attach'] = @($attach)
        $data['sent'] = $false
        return $data
    }

    try {
        if ($null -ne $newBody) {
            if ($asHtml) {
                $item.HTMLBody = $newBody
            } else {
                $item.BodyFormat = $script:OL_FORMAT_PLAIN
                $item.Body = $newBody
            }
        }
        if (-not [string]::IsNullOrEmpty($subject)) { $item.Subject = $subject }
        if (-not [string]::IsNullOrEmpty($from)) { [void](Set-OlDraftSender -Ns $Ns -Mail $item -From $from) }
    } catch {
        if ($null -ne $script:OlErrorPayload) { throw }
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not update the draft'
    }
    foreach ($p in $attach) {
        $winPath = [string]$p
        if ([string]::IsNullOrEmpty($winPath)) { continue }
        try {
            [void]$item.Attachments.Add($winPath)
        } catch {
            Invoke-OlFail -Code 'bad_argument' -Detail $_.Exception.Message `
                -Message ("could not attach '" + $winPath + "'. It must be a Windows path reachable from Outlook.")
        }
    }
    if ($hasText) { Set-OlDraftTextMark -Mail $item -Text $text }
    try {
        $item.Save()
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not save the draft'
    }

    $data = @{}
    $data['draft_ref'] = Get-Safe { Get-OlRef -Item $item } $ref
    $data['subject'] = Get-Safe { $item.Subject } $null
    $data['body'] = $mode
    $data['saved_to'] = Get-Safe { Get-OlFolderPath -Folder $item.Parent } 'Drafts'
    $data['sent'] = $false
    $data['dry_run'] = $false
    return $data
}

# ----------------------------------------------------------- rules & folders
#
# 0.3.4 additions: create folders and manage server-side receive rules, so
# machine mail (expense tools, ERP workflow failures, helpdesk approvals, M365
# platform digests, vendor marketing) can be filed automatically. Rules live per store.
# Conditions/Actions are addressed by their documented OlRuleConditionType /
# OlRuleActionType numbers: subject=2, sender_address=17; move_to_folder=1,
# mark_read=19, stop=21.

$script:OL_CONDITION_NAMES = @{
    1 = 'from'; 2 = 'subject'; 3 = 'account'; 4 = 'only_to_me'; 5 = 'to';
    6 = 'importance'; 7 = 'sensitivity'; 8 = 'flagged_for_action'; 9 = 'cc';
    10 = 'to_or_cc'; 11 = 'not_to'; 12 = 'sent_to'; 13 = 'body';
    14 = 'body_or_subject'; 15 = 'message_header'; 16 = 'recipient_address';
    17 = 'sender_address'; 18 = 'category'; 19 = 'oof'; 20 = 'has_attachment';
    21 = 'size_range'; 22 = 'date_range'; 23 = 'form_name'; 24 = 'property';
    25 = 'sender_in_address_book'; 26 = 'meeting_invite_or_update';
    27 = 'local_machine_only'; 28 = 'other_machine'; 29 = 'any_category';
    30 = 'from_rss_feed'; 31 = 'from_any_rss_feed'
}
$script:OL_ACTION_NAMES = @{
    1 = 'move_to_folder'; 2 = 'assign_to_category'; 3 = 'delete';
    4 = 'delete_permanently'; 5 = 'copy_to_folder'; 6 = 'forward';
    7 = 'forward_as_attachment'; 8 = 'redirect'; 9 = 'server_reply';
    10 = 'template'; 11 = 'flag_for_action_in_days'; 12 = 'flag_color';
    13 = 'flag_clear'; 14 = 'importance'; 15 = 'sensitivity'; 16 = 'print';
    17 = 'play_sound'; 18 = 'start_application'; 19 = 'mark_read';
    20 = 'run_script'; 21 = 'stop'; 22 = 'custom_action'; 23 = 'new_item_alert';
    24 = 'desktop_alert'; 25 = 'notify_read'; 26 = 'notify_delivery';
    27 = 'cc_message'; 28 = 'defer'; 41 = 'mark_as_task'
}

function Get-OlStoreObject {
    param($Ns, [string]$Name)
    if ([string]::IsNullOrEmpty($Name)) {
        try { return $Ns.DefaultStore } catch { return $null }
    }
    $key = $Name.Trim().ToLower()
    $n = 0
    try { $n = [int]$Ns.Stores.Count } catch { $n = 0 }
    for ($i = 1; $i -le $n; $i++) {
        $st = $null
        try { $st = $Ns.Stores.Item($i) } catch { $st = $null }
        if ($null -eq $st) { continue }
        $nm1 = $null
        try { $nm1 = [string]$st.DisplayName } catch { $nm1 = $null }
        $nm2 = $null
        try { $nm2 = [string]$st.GetRootFolder().Name } catch { $nm2 = $null }
        foreach ($nm in @($nm1, $nm2)) {
            if ((-not [string]::IsNullOrEmpty($nm)) -and ($nm.Trim().ToLower() -eq $key)) { return $st }
        }
    }
    return $null
}

function Get-OlRuleSummary {
    param($Rule)
    $conds = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($c in @($Rule.Conditions)) {
            $enabled = $false
            try { $enabled = [bool]$c.Enabled } catch { }
            if (-not $enabled) { continue }
            $h = @{}
            $ct = 0
            try { $ct = [int]$c.ConditionType } catch { $ct = 0 }
            $h['type'] = $ct
            if ($script:OL_CONDITION_NAMES.ContainsKey($ct)) { $h['type_name'] = $script:OL_CONDITION_NAMES[$ct] }
            if (@(2, 13, 14, 15, 16, 17) -contains $ct) {
                $words = $null
                try { $words = @($c.Text) } catch { $words = $null }
                if (($null -ne $words) -and ($words.Count -gt 0)) {
                    $clean = @($words | Where-Object { $null -ne $_ })
                    if ($clean.Count -gt 0) { $h['text'] = $clean }
                }
            }
            if ($ct -eq 1) {
                try {
                    $rcpts = New-Object System.Collections.Generic.List[object]
                    foreach ($r in @($c.Recipients)) {
                        $rh = @{}
                        $rh['address'] = Get-Safe { [string]$r.Address } $null
                        $rh['name'] = Get-Safe { [string]$r.Name } $null
                        $rh['resolved'] = Get-Safe { [bool]$r.Resolved } $false
                        [void]$rcpts.Add($rh)
                    }
                    if ($rcpts.Count -gt 0) { $h['recipients'] = @($rcpts.ToArray()) }
                } catch { }
            }
            [void]$conds.Add($h)
        }
    } catch { }
    $acts = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($ac in @($Rule.Actions)) {
            $enabled = $false
            try { $enabled = [bool]$ac.Enabled } catch { }
            if (-not $enabled) { continue }
            $h = @{}
            $at = 0
            try { $at = [int]$ac.ActionType } catch { $at = 0 }
            $h['type'] = $at
            if ($script:OL_ACTION_NAMES.ContainsKey($at)) { $h['type_name'] = $script:OL_ACTION_NAMES[$at] }
            if ($at -eq 1) {
                $h['folder_path'] = Get-Safe { [string]$ac.Folder.FolderPath } $null
            }
            [void]$acts.Add($h)
        }
    } catch { }
    $h = @{}
    $h['name'] = Get-Safe { [string]$Rule.Name } $null
    $h['enabled'] = Get-Safe { [bool]$Rule.Enabled } $false
    $h['rule_type'] = Get-Safe { [int]$Rule.RuleType } $null
    $h['execution_order'] = Get-Safe { [int]$Rule.ExecutionOrder } $null
    $h['is_client_rule'] = Get-Safe { [bool]$Rule.IsClientRule } $null
    $h['conditions'] = @($conds.ToArray())
    $h['actions'] = @($acts.ToArray())
    return $h
}

function Invoke-OlRulesList {
    param($Ol, $Ns, $OpArgs)
    $store = Get-OlStoreObject -Ns $Ns -Name (Get-OlString $OpArgs 'store' $null)
    if ($null -eq $store) {
        Invoke-OlFail -Code 'not_found' -Extra @{ available_stores = (Get-OlStoreNames $Ns) } `
            -Message 'no store like that (omit --store for the default store)'
    }
    $rules = $null
    try { $rules = $store.GetRules() } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not read the rules of this store'
    }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($rules)) {
        [void]$out.Add((Get-OlRuleSummary -Rule $r))
    }
    $data = @{}
    $data['store'] = Get-Safe { [string]$store.DisplayName } $null
    $data['count'] = $out.Count
    $data['rules'] = @($out.ToArray())
    return $data
}

function Get-OlConditionByType {
    # The Conditions/Actions collections index positionally, not by the
    # documented enum value -- Item(17) returns whatever sits at position 17.
    # Enumerate and match on the type property instead.
    param($Rule, [int]$Type)
    try {
        foreach ($c in @($Rule.Conditions)) {
            $ct = -1
            try { $ct = [int]$c.ConditionType } catch { $ct = -1 }
            if ($ct -eq $Type) { return $c }
        }
    } catch { }
    return $null
}

function Get-OlActionByType {
    param($Rule, [int]$Type)
    try {
        foreach ($a in @($Rule.Actions)) {
            $at = -1
            try { $at = [int]$a.ActionType } catch { $at = -1 }
            if ($at -eq $Type) { return $a }
        }
    } catch { }
    return $null
}

function Invoke-OlRulesAdd {
    param($Ol, $Ns, $OpArgs)

    $name = Get-OlString $OpArgs 'name' $null
    if ([string]::IsNullOrEmpty($name)) { Invoke-OlFail -Code 'bad_argument' -Message 'rules-add needs --name' }
    $senderWords = Get-OlArgArray $OpArgs 'sender_contains'
    $subjectWords = Get-OlArgArray $OpArgs 'subject_contains'
    $headerWords = Get-OlArgArray $OpArgs 'header_contains'
    $fromList = Get-OlArgArray $OpArgs 'from'
    $moveTo = Get-OlString $OpArgs 'move_to' $null
    $shared = Get-OlString $OpArgs 'shared' $null
    $markRead = Get-OlBool $OpArgs 'mark_read' $false
    $stopHere = Get-OlBool $OpArgs 'stop' $false
    $enabled = Get-OlBool $OpArgs 'enabled' $true
    $replace = Get-OlBool $OpArgs 'replace' $false
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    if (($senderWords.Count -eq 0) -and ($fromList.Count -eq 0) -and ($subjectWords.Count -eq 0) -and ($headerWords.Count -eq 0)) {
        Invoke-OlFail -Code 'bad_argument' -Message 'rules-add needs --from, --subject-contains or --header-contains'
    }
    if ($senderWords.Count -gt 0) {
        Invoke-OlFail -Code 'bad_argument' `
            -Message ('--sender-contains is not settable through the Outlook object model on this build (the condition arrives as a base-class object). Use --from with exact addresses, or --header-contains with the domain, which also matches rotating sender addresses.')
    }
    if ([string]::IsNullOrEmpty($moveTo)) { $moveTo = $null }  # optional: a rule may only mark read

    $store = Get-OlStoreObject -Ns $Ns -Name (Get-OlString $OpArgs 'store' $null)
    if ($null -eq $store) {
        Invoke-OlFail -Code 'not_found' -Extra @{ available_stores = (Get-OlStoreNames $Ns) } `
            -Message 'no store like that (omit --store for the default store)'
    }

    # Resolving the target is read-only and validates the path even in a dry run.
    $folder = $null
    $movePath = $null
    if ($null -ne $moveTo) {
        $folder = Resolve-OlFolder -Ns $Ns -Path $moveTo -Shared $shared
        $movePath = Get-Safe { [string]$folder.FolderPath } $moveTo
    }

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['name'] = $name
        if ($senderWords.Count -gt 0) { $data['sender_contains'] = @($senderWords) }
        if ($subjectWords.Count -gt 0) { $data['subject_contains'] = @($subjectWords) }
        if ($headerWords.Count -gt 0) { $data['header_contains'] = @($headerWords) }
        if ($fromList.Count -gt 0) { $data['from'] = @($fromList) }
        if ($null -ne $movePath) { $data['move_to'] = $movePath }
        $data['mark_read'] = [bool]$markRead
        $data['stop'] = [bool]$stopHere
        $data['enabled'] = [bool]$enabled
        return $data
    }

    $rules = $null
    try { $rules = $store.GetRules() } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not read the rules of this store'
    }

    $existing = $null
    foreach ($r in @($rules)) {
        $rn = Get-Safe { [string]$r.Name } ''
        if ((-not [string]::IsNullOrEmpty($rn)) -and ($rn.Trim().ToLower() -eq $name.Trim().ToLower())) { $existing = $r; break }
    }
    $replaced = $false
    if ($null -ne $existing) {
        if (-not $replace) {
            Invoke-OlFail -Code 'rule_exists' -Message ("a rule named '" + $name + "' already exists; pass --replace to overwrite it")
        }
        try { $rules.Remove((Get-Safe { [int]$existing.ExecutionOrder } 1)) } catch {
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not remove the existing rule'
        }
        $replaced = $true
    }

    # 0 = olRuleReceive
    $rule = $null
    try { $rule = $rules.Create($name, 0) } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not create the rule'
    }

    if ($senderWords.Count -gt 0) {
        # Typed access via the named property is the documented way; the
        # enumerated object of the same type can arrive as a base-class wrapper.
        $cond = $null
        try {
            try { $cond = $rule.Conditions.SenderAddress } catch { $cond = $null }
            if ($null -eq $cond) { $cond = Get-OlConditionByType -Rule $rule -Type 17 }
            if ($null -eq $cond) { throw 'sender-address condition not available on this rule type' }
            try { $cond.Text = [string[]]@($senderWords) } catch { $cond.Text = [object[]]@($senderWords) }
            $cond.Enabled = $true
        } catch {
            $dbg = @{ }
            try { $dbg['condition_type'] = [int]$cond.ConditionType } catch { }
            try { $dbg['members'] = @($cond | Get-Member -MemberType Property | ForEach-Object { $_.Name }) } catch { }
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
                -Message 'could not set the sender-address condition' -Extra @{ condition_debug = $dbg }
        }
    }
    if ($headerWords.Count -gt 0) {
        $cond = $null
        try {
            try { $cond = $rule.Conditions.MessageHeader } catch { $cond = $null }
            if ($null -eq $cond) { $cond = Get-OlConditionByType -Rule $rule -Type 15 }
            if ($null -eq $cond) { throw 'message-header condition not available on this rule type' }
            try { $cond.Text = [string[]]@($headerWords) } catch { $cond.Text = [object[]]@($headerWords) }
            $cond.Enabled = $true
        } catch {
            $dbg = @{ }
            try { $dbg['condition_type'] = [int]$cond.ConditionType } catch { }
            try { $dbg['members'] = @($cond | Get-Member -MemberType Property | ForEach-Object { $_.Name }) } catch { }
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
                -Message 'could not set the message-header condition' -Extra @{ condition_debug = $dbg }
        }
    }
    if ($subjectWords.Count -gt 0) {
        $cond = $null
        try {
            try { $cond = $rule.Conditions.Subject } catch { $cond = $null }
            if ($null -eq $cond) { $cond = Get-OlConditionByType -Rule $rule -Type 2 }
            if ($null -eq $cond) { throw 'subject condition not available on this rule type' }
            try { $cond.Text = [string[]]@($subjectWords) } catch { $cond.Text = [object[]]@($subjectWords) }
            $cond.Enabled = $true
        } catch {
            $dbg = @{ }
            try { $dbg['condition_type'] = [int]$cond.ConditionType } catch { }
            try { $dbg['members'] = @($cond | Get-Member -MemberType Property | ForEach-Object { $_.Name }) } catch { }
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
                -Message 'could not set the subject condition' -Extra @{ condition_debug = $dbg }
        }
    }
    if ($fromList.Count -gt 0) {
        try {
            $cond = Get-OlConditionByType -Rule $rule -Type 1    # olConditionFrom
            if ($null -eq $cond) { throw 'from condition not available on this rule type' }
            foreach ($addr in $fromList) {
                $rcpt = $cond.Recipients.Add([string]$addr)
                try { [void]$rcpt.Resolve() } catch { }
            }
            $cond.Enabled = $true
        } catch {
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not set the from condition'
        }
    }

    $act = $null
    $moveState = $null
    if ($null -ne $folder) {
        try {
            # The named property can hand back a detached copy whose Folder
            # assignment silently does not stick; the enumerated object is the
            # live one. Try the enumerated item first here.
            $act = Get-OlActionByType -Rule $rule -Type 1
            if ($null -eq $act) { try { $act = $rule.Actions.MoveToFolder } catch { $act = $null } }
            if ($null -eq $act) { throw 'move-to-folder action not available on this rule type' }
            $act.Folder = $folder
            $check = Get-Safe { [string]$act.Folder.FolderPath } ''
            if ([string]::IsNullOrEmpty($check)) {
                # PowerShell can silently no-op a property put whose value is
                # itself a COM object. Put it late-bound through IDispatch.
                try {
                    [void][System.__ComObject].InvokeMember('Folder', [System.Reflection.BindingFlags]::SetProperty, $null, $act, @($folder))
                } catch { }
                $check = Get-Safe { [string]$act.Folder.FolderPath } ''
            }
            if ([string]::IsNullOrEmpty($check)) {
                # Last variant: a freshly fetched folder handle.
                $fid = Get-Safe { [string]$folder.EntryID } $null
                $sid = Get-Safe { [string]$folder.StoreID } $null
                $fresh = $null
                if (-not [string]::IsNullOrEmpty($fid)) {
                    try { $fresh = $Ns.GetFolderFromID($fid, $sid) } catch { $fresh = $null }
                }
                if ($null -ne $fresh) {
                    try {
                        [void][System.__ComObject].InvokeMember('Folder', [System.Reflection.BindingFlags]::SetProperty, $null, $act, @($fresh))
                    } catch { }
                    $check = Get-Safe { [string]$act.Folder.FolderPath } ''
                }
            }
            if ([string]::IsNullOrEmpty($check)) { throw ('folder assignment did not stick on the move action (FolderPath empty after three put attempts)') }
            $act.Enabled = $true
            $moveState = @{}
            try { $moveState['enabled'] = [bool]$act.Enabled } catch { }
            try { $moveState['folder'] = Get-Safe { [string]$act.Folder.FolderPath } $null } catch { }
            try { $moveState['members'] = @($act | Get-Member -MemberType Property | ForEach-Object { $_.Name }) } catch { }
        } catch {
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not set the move action'
        }
    }
    if ($stopHere) {
        try {
            try { $act = $rule.Actions.Stop } catch { $act = $null }
            if ($null -eq $act) { $act = Get-OlActionByType -Rule $rule -Type 21 }
            if ($null -eq $act) { throw 'stop action not available on this rule type' }
            $act.Enabled = $true
        } catch {
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not set the stop action'
        }
    }

    try { $rule.Enabled = [bool]$enabled } catch { }

    try { [void]$rules.Save() } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not save the rules' -Extra @{ move_state = $moveState }
    }

    # Mark-as-read cannot be enabled on the fresh in-memory rule on this build
    # (Outlook rejects it even alone). A saved rule sometimes accepts it in a
    # second pass -- try that, but never fail the rule over it.
    $markReadApplied = $false
    if ($markRead) {
        try {
            $rules3 = $store.GetRules()
            $saved3 = $null
            foreach ($r in @($rules3)) {
                $rn = Get-Safe { [string]$r.Name } ''
                if ((-not [string]::IsNullOrEmpty($rn)) -and ($rn.Trim().ToLower() -eq $name.Trim().ToLower())) { $saved3 = $r; break }
            }
            if ($null -ne $saved3) {
                $act = $null
                try { $act = $saved3.Actions.MarkAsRead } catch { $act = $null }
                if ($null -eq $act) { $act = Get-OlActionByType -Rule $saved3 -Type 19 }
                if ($null -ne $act) {
                    $act.Enabled = $true
                    [void]$rules3.Save()
                    $markReadApplied = $true
                }
            }
        } catch { $markReadApplied = $false }
    }

    # Read back what actually persisted and report that, not the intent.
    $saved = $null
    try {
        $rules2 = $store.GetRules()
        foreach ($r in @($rules2)) {
            $rn = Get-Safe { [string]$r.Name } ''
            if ((-not [string]::IsNullOrEmpty($rn)) -and ($rn.Trim().ToLower() -eq $name.Trim().ToLower())) { $saved = $r; break }
        }
    } catch { }

    $data = @{}
    $data['name'] = $name
    $data['replaced'] = [bool]$replaced
    if ($null -ne $movePath) { $data['move_to'] = $movePath }
    $data['mark_read_applied'] = [bool]$markReadApplied
    $data['dry_run'] = $false
    if ($null -ne $saved) {
        $summary = Get-OlRuleSummary -Rule $saved
        foreach ($k in @($summary.Keys)) { $data[$k] = $summary[$k] }
    }
    return $data
}

function Invoke-OlRulesDelete {
    param($Ol, $Ns, $OpArgs)

    $name = Get-OlString $OpArgs 'name' $null
    if ([string]::IsNullOrEmpty($name)) { Invoke-OlFail -Code 'bad_argument' -Message 'rules-delete needs --name' }

    $store = Get-OlStoreObject -Ns $Ns -Name (Get-OlString $OpArgs 'store' $null)
    if ($null -eq $store) {
        Invoke-OlFail -Code 'not_found' -Extra @{ available_stores = (Get-OlStoreNames $Ns) } `
            -Message 'no store like that (omit --store for the default store)'
    }

    $rules = $null
    try { $rules = $store.GetRules() } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not read the rules of this store'
    }
    $existing = $null
    foreach ($r in @($rules)) {
        $rn = Get-Safe { [string]$r.Name } ''
        if ((-not [string]::IsNullOrEmpty($rn)) -and ($rn.Trim().ToLower() -eq $name.Trim().ToLower())) { $existing = $r; break }
    }
    if ($null -eq $existing) {
        Invoke-OlFail -Code 'not_found' -Message ("no rule named '" + $name + "'")
    }
    $order = Get-Safe { [int]$existing.ExecutionOrder } 0
    try { $rules.Remove($order) } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not remove the rule'
    }
    try { [void]$rules.Save() } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not save the rules after removal'
    }

    $data = @{}
    $data['name'] = $name
    $data['deleted'] = $true
    $data['execution_order'] = $order
    $data['dry_run'] = $false
    return $data
}

function Get-OlFolderStat {
    param($Folder)

    # Same source as the folders command: UnReadItemCount and Items.Count.
    $stat = @{ total = $null; unread = $null }
    $stat['total'] = Get-Safe { [int]$Folder.Items.Count } $null
    $stat['unread'] = Get-Safe { [int]$Folder.UnReadItemCount } $null
    return $stat
}

# Run Rules Now, programmatically: Rule.Execute() applies a rule's actions to
# the items of one folder, which is what the wizard's "Run Rules Now" button
# does. Execute() gives no feedback, so the command wraps the run with the
# unread counts of the target folder before and after. Only enabled receive
# rules run, the same set the wizard offers; --all runs them in ExecutionOrder.
function Invoke-OlRulesRun {
    param($Ol, $Ns, $OpArgs)

    $name = Get-OlString $OpArgs 'name' $null
    $runAll = Get-OlBool $OpArgs 'all' $false
    if (([string]::IsNullOrEmpty($name)) -and (-not $runAll)) {
        Invoke-OlFail -Code 'bad_argument' -Message 'rules-run needs --name or --all'
    }

    $store = Get-OlStoreObject -Ns $Ns -Name (Get-OlString $OpArgs 'store' $null)
    if ($null -eq $store) {
        Invoke-OlFail -Code 'not_found' -Extra @{ available_stores = (Get-OlStoreNames $Ns) } `
            -Message 'no store like that (omit --store for the default store)'
    }

    $rules = $null
    try { $rules = $store.GetRules() } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not read the rules of this store'
    }

    $candidates = New-Object System.Collections.Generic.List[object]
    if ($runAll) {
        foreach ($r in @($rules)) {
            $enabled = Get-Safe { [bool]$r.Enabled } $false
            $rtype = Get-Safe { [int]$r.RuleType } -1
            if ($enabled -and ($rtype -eq 0)) { [void]$candidates.Add($r) }
        }
    } else {
        $existing = $null
        foreach ($r in @($rules)) {
            $rn = Get-Safe { [string]$r.Name } ''
            if ((-not [string]::IsNullOrEmpty($rn)) -and ($rn.Trim().ToLower() -eq $name.Trim().ToLower())) { $existing = $r; break }
        }
        if ($null -eq $existing) {
            Invoke-OlFail -Code 'not_found' -Message ("no rule named '" + $name + "'")
        }
        if (-not (Get-Safe { [bool]$existing.Enabled } $false)) {
            Invoke-OlFail -Code 'bad_argument' -Message ("rule '" + $name + "' is disabled")
        }
        [void]$candidates.Add($existing)
    }
    $candidates = @($candidates | Sort-Object { (Get-Safe { [int]$_.ExecutionOrder } 0) })

    # Default target is the store's Inbox; --folder overrides, and
    # --include-subfolders widens the sweep the way the wizard checkbox does.
    $folderPath = Get-OlString $OpArgs 'folder' $null
    if ([string]::IsNullOrEmpty($folderPath)) {
        try { $target = $store.GetDefaultFolder(6) } catch {   # olFolderInbox
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not resolve the Inbox of this store'
        }
    } else {
        $target = Resolve-OlFolder -Ns $Ns -Path $folderPath -Shared (Get-OlString $OpArgs 'shared' $null)
    }
    $targetPath = Get-Safe { [string]$target.FolderPath } $null
    $includeSub = Get-OlBool $OpArgs 'include_subfolders' $false

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($r in $candidates) { [void]$names.Add((Get-Safe { [string]$r.Name } '?')) }

    $data = @{}
    $data['store'] = Get-Safe { [string]$store.FolderPath } $null
    $data['folder'] = $targetPath
    $data['include_subfolders'] = [bool]$includeSub
    $data['rules'] = @($names.ToArray())
    $data['folder_stats_before'] = Get-OlFolderStat $target

    if (Get-OlBool $OpArgs 'dry_run' $false) {
        $data['dry_run'] = $true
        $data['note'] = 'dry run: nothing was executed'
        return $data
    }

    $executed = New-Object System.Collections.Generic.List[string]
    foreach ($r in $candidates) {
        $rn = Get-Safe { [string]$r.Name } '?'
        try { [void]$r.Execute($false, $target, $includeSub) } catch {
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
                -Message ("rule '" + $rn + "' could not execute")
        }
        [void]$executed.Add($rn)
    }

    $data['executed'] = @($executed.ToArray())
    $data['folder_stats_after'] = Get-OlFolderStat $target
    $uBefore = $data['folder_stats_before']['unread']
    $uAfter = $data['folder_stats_after']['unread']
    if (($null -ne $uBefore) -and ($null -ne $uAfter)) {
        $data['unread_delta'] = ($uAfter - $uBefore)
    }
    $data['dry_run'] = $false
    return $data
}

function Invoke-OlFolderAdd {
    param($Ol, $Ns, $OpArgs)

    $parentPath = Get-OlString $OpArgs 'folder' $null
    $nameChain = Get-OlArgArray $OpArgs 'name'
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    if ([string]::IsNullOrEmpty($parentPath)) {
        Invoke-OlFail -Code 'bad_argument' -Message 'folder-add needs --folder (the parent path)'
    }
    if ($nameChain.Count -eq 0) {
        Invoke-OlFail -Code 'bad_argument' -Message 'folder-add needs --name'
    }

    $parent = Resolve-OlFolder -Ns $Ns -Path $parentPath -Shared (Get-OlString $OpArgs 'shared' $null)
    $parentResolved = Get-Safe { [string]$parent.FolderPath } $parentPath

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['parent'] = $parentResolved
        $data['create'] = @($nameChain)
        return $data
    }

    $cur = $parent
    $created = New-Object System.Collections.Generic.List[object]
    $existed = New-Object System.Collections.Generic.List[object]
    foreach ($seg in $nameChain) {
        $seg = [string]$seg
        if ([string]::IsNullOrEmpty($seg)) { continue }
        $child = Find-OlChild $cur $seg
        if ($null -eq $child) {
            try { $child = $cur.Folders.Add($seg) } catch {
                Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
                    -Message ("could not create folder '" + $seg + "' under '" + (Get-Safe { [string]$cur.Name } '?') + "'")
            }
            [void]$created.Add($seg)
        } else {
            [void]$existed.Add($seg)
        }
        $cur = $child
    }

    $data = @{}
    $data['parent'] = $parentResolved
    $data['created'] = @($created.ToArray())
    $data['already_existed'] = @($existed.ToArray())
    $data['folder_path'] = Get-Safe { [string]$cur.FolderPath } $null
    $data['dry_run'] = $false
    return $data
}

# ------------------------------------------------------------------ dispatch

$script:OL_OPS = @('doctor', 'folders', 'list', 'get', 'thread', 'save_attachments', 'move',
    'mark', 'flag', 'categorize', 'draft_reply', 'draft_new', 'draft_update',
    'folder_add', 'rules_list', 'rules_add', 'rules_delete', 'rules_run')

# $Conn lets batch attach to Outlook once for many ops.
function Invoke-OlOp {
    param([string]$Op, $OpArgs, $Conn = $null)

    # Validated before touching COM, so an unknown op always reports
    # bad_argument rather than whatever the COM attach happens to say.
    if (-not ($script:OL_OPS -contains $Op)) {
        Invoke-OlFail -Code 'bad_argument' -Message ("unknown op '" + [string]$Op + "'. Known ops: " +
            [string]::Join(', ', $script:OL_OPS) + '.')
    }

    if ($null -eq $Conn) { $Conn = Connect-Outlook }
    $ol = $Conn['app']
    $ns = $Conn['ns']

    $data = $null
    switch ($Op) {
        'doctor' { $data = Invoke-OlDoctor -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'folders' { $data = Invoke-OlFolders -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'list' { $data = Invoke-OlList -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'get' { $data = Invoke-OlGet -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'thread' { $data = Invoke-OlThread -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'save_attachments' { $data = Invoke-OlSaveAttachments -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'move' { $data = Invoke-OlMove -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'mark' { $data = Invoke-OlMark -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'flag' { $data = Invoke-OlFlag -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'categorize' { $data = Invoke-OlCategorize -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'draft_reply' { $data = Invoke-OlDraftReply -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'draft_new' { $data = Invoke-OlDraftNew -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'draft_update' { $data = Invoke-OlDraftUpdate -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'folder_add' { $data = Invoke-OlFolderAdd -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'rules_list' { $data = Invoke-OlRulesList -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'rules_add' { $data = Invoke-OlRulesAdd -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'rules_delete' { $data = Invoke-OlRulesDelete -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'rules_run' { $data = Invoke-OlRulesRun -Ol $ol -Ns $ns -OpArgs $OpArgs }
        default {
            Invoke-OlFail -Code 'bad_argument' -Message ("unknown op '" + [string]$Op + "'. Known ops: " +
                [string]::Join(', ', $script:OL_OPS) + '.')
        }
    }
    if ($null -eq $data) { $data = @{} }
    return $data
}

# ============================================================== CLI front-end
#
# Everything above this line is the COM engine. Everything below turns a normal
# command line into one engine call, and is where the concerns that do not need
# COM live: option parsing, date arithmetic, the config file and the audit log.

# 0.3.4 adds folder-add / rules-list / rules-add (see the section above the
# dispatch table) and the [string[]] date-format fix for --since/--until.
# 0.3.5 adds rules-run (Run Rules Now on a folder, with before/after counts).
# 0.4.0 adds thread, draft-update, batch, list paging/multi-folder/body search/
# exclusions/marks, draft --from/--save-to/--supersede, get --strip-quoted,
# doctor --quick and --json-out.
# 0.4.1: draft-update matches its text ignoring whitespace; text drafts are
# plain format (no mailto duplication); get shows a draft's sender fields.
# 0.5.0: guide (bundled agent guide), machine policy (policy.json), per-user
# installer with bundled skills.
$script:OLCTL_VERSION = '0.5.0'

$script:USAGE = @'
olctl <command> [options]     local Outlook Classic CLI -- one JSON object per run

Start here
  guide                                    how to use olctl well (plain text);
                                           AI agents: read this first

Reading
  doctor [--quick]                         connection, stores, policy, protected folders
  folders [--folder P] [--depth N]         folder tree with counts
  list --folder P [--folder P2 ...] [filters]   messages, newest first
  get REF [--body text|html|none] [--strip-quoted] [--headers]
  thread REF [--also P ...] [--since T] [--body text|none] [--keep-quoted]
                                           the whole conversation, oldest first
  save-attachments REF --out DIR [--pattern "*.pdf"]

Changing mail
  move REF --to P                          move between folders
  mark REF read|unread                     read state
  flag REF [--text T] [--clear]            follow-up flag
  categorize REF [--add C] [--remove C] [--set A B ...]
  batch [--file F | -] [--stop-on-error]   many commands, one Outlook attach;
                                           one JSON array of argv per line

Drafts (never sent)
  draft-reply REF --text ... [--all] [--html] [--subject S] [--attach F]
        [--from ADDR] [--save-to P] [--supersede OLD_DRAFT_REF]
  draft-new --to X --subject S --text ... [--cc Y] [--html] [--attach F]
        [--from ADDR] [--save-to P]
  draft-update DRAFT_REF [--text ...] [--whole-body] [--html] [--subject S]
        [--from ADDR] [--attach F]

Folders and rules
  folder-add --folder P --name A[/B]       create folders (a chain of names)
  rules-list [--store S]                   inbox rules of a store (default store)
  rules-add --name N --move-to P           new receive rule
        [--from a@b,c@d] [--subject-contains w1,w2] [--header-contains d1,d2]
        [--mark-read] [--stop] [--disabled] [--replace]
  rules-delete --name N [--store S]        delete a rule by name
  rules-run [--name N | --all]             run rules now on a folder (Inbox
        [--folder P] [--include-subfolders] by default; reports unread counts)

list filters
  --unread --read --flagged --has-attachments
  --since 30m|6h|3d|2w|YYYY-MM-DD  --until ...
  --since-mark NAME  --save-mark NAME      remember where the last scan ended
  --sender SUB  --subject SUB  --category C  --body-contains TEXT
  --exclude-sender SUB  --exclude-subject SUB  --preset NAME   (repeatable)
  --include-subfolders  --after-cursor C   (C = next_cursor of the last page)
  --limit N (25)  --scan-max N (2000)  --oldest-first
  --no-preview  --any-class  --count-only

Global
  --pretty        indent the JSON
  --dry-run       mutating commands report the plan and change nothing
  --json-out F    write the JSON to F as UTF-8, print only where it went
  --version       print the version
  --help          this text

Folder paths: "Inbox", "Inbox/Vendors", or store-qualified "projects/inbox".
The English tokens inbox/drafts/sent items/deleted items/junk email/outbox work
under any store name and resolve inside that store, so localised mailboxes
(Posteingang, Dorucena posta) need no special handling. "archive" is NOT a
token: use the folder's real name, e.g. "projects/Archive".
'@

# ------------------------------------------------------------------- settings

function Get-OlctlHome {
    if (-not [string]::IsNullOrEmpty($env:OLCTL_HOME)) { return $env:OLCTL_HOME }
    $base = $env:USERPROFILE
    if ([string]::IsNullOrEmpty($base)) { $base = $env:HOME }
    return (Join-Path $base '.olctl')
}

<#
    Machine policy: %ProgramData%\olctl\policy.json, deployed by IT.
        { "read_only": true, "disabled_commands": ["rules-add", "rules-delete"] }
    read_only blocks every command that changes mail, drafts, folders or rules
    (their --dry-run still works); disabled_commands blocks those outright.
    OLCTL_POLICY_PATH names an extra policy file that can only tighten: flags
    are OR-ed and lists unioned, so it is safe for tests and cannot loosen
    the machine policy. A policy file that exists but cannot be read fails
    closed. This is a guardrail for an assistant, not a security boundary:
    a per-user install is writable by its user.
#>
$script:OLCTL_MUTATING = @('move', 'mark', 'flag', 'categorize', 'draft-reply', 'draft-new',
    'draft-update', 'folder-add', 'rules-add', 'rules-delete', 'rules-run')

function Read-OlctlPolicy {
    $pol = @{ read_only = $false; disabled_commands = @(); sources = @() }
    $paths = New-Object System.Collections.Generic.List[object]
    if (-not [string]::IsNullOrEmpty($env:ProgramData)) {
        [void]$paths.Add((Join-Path (Join-Path $env:ProgramData 'olctl') 'policy.json'))
    }
    if (-not [string]::IsNullOrEmpty($env:OLCTL_POLICY_PATH)) { [void]$paths.Add($env:OLCTL_POLICY_PATH) }
    $disabled = New-Object System.Collections.Generic.List[object]
    $sources = New-Object System.Collections.Generic.List[object]
    foreach ($path in $paths) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $parsed = $null
        try {
            $parsed = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($path, [Text.Encoding]::UTF8))
        } catch {
            Invoke-OlFail -Code 'bad_policy' -Detail $_.Exception.Message `
                -Message ('the machine policy ' + $path + ' cannot be read, so olctl refuses to run. Ask IT to fix or remove it.')
        }
        [void]$sources.Add($path)
        if ($null -eq $parsed) { continue }
        $ro = $parsed.PSObject.Properties['read_only']
        if (($null -ne $ro) -and ($ro.Value -eq $true)) { $pol['read_only'] = $true }
        $dc = $parsed.PSObject.Properties['disabled_commands']
        if ($null -ne $dc) {
            foreach ($c in @($dc.Value)) {
                $name = ([string]$c).Trim().ToLowerInvariant()
                if (($name.Length -gt 0) -and (-not $disabled.Contains($name))) { [void]$disabled.Add($name) }
            }
        }
    }
    $pol['disabled_commands'] = @($disabled.ToArray())
    $pol['sources'] = @($sources.ToArray())
    return $pol
}

function Assert-OlctlPolicy {
    param([string]$Cmd, $Policy, [bool]$DryRun)
    if ($null -eq $Policy) { return }
    $extra = @{ policy = @($Policy['sources']) }
    if (@($Policy['disabled_commands']) -contains $Cmd) {
        Invoke-OlFail -Code 'disabled_by_policy' -Extra $extra `
            -Message ("'" + $Cmd + "' is disabled on this machine by policy")
    }
    if ($Policy['read_only'] -and ($script:OLCTL_MUTATING -contains $Cmd) -and (-not $DryRun)) {
        Invoke-OlFail -Code 'disabled_by_policy' -Extra $extra `
            -Message ("olctl is read-only on this machine by policy; '" + $Cmd + "' can only run with --dry-run")
    }
}

# The guide ships next to the script (guide\GUIDE.md); the active policy is
# appended so an agent reading it knows what it may do on this machine.
function Get-OlctlGuide {
    param($Policy)
    $path = Join-Path (Join-Path $PSScriptRoot 'guide') 'GUIDE.md'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        Invoke-OlFail -Code 'bad_install' -Message ('guide not found at ' + $path + '; reinstall olctl')
    }
    $text = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8).TrimEnd()
    $lines = New-Object System.Collections.Generic.List[object]
    [void]$lines.Add('')
    [void]$lines.Add('## On this machine')
    [void]$lines.Add('')
    [void]$lines.Add('- olctl version: ' + $script:OLCTL_VERSION)
    if ($Policy['read_only']) {
        [void]$lines.Add('- Policy: READ-ONLY. Commands that change mail or drafts only work with --dry-run.')
    } else {
        [void]$lines.Add('- Policy: drafts and mail changes allowed (still: ask the user first).')
    }
    if (@($Policy['disabled_commands']).Count -gt 0) {
        [void]$lines.Add('- Disabled by policy: ' + [string]::Join(', ', @($Policy['disabled_commands'])))
    }
    return ($text + "`n" + [string]::Join("`n", $lines.ToArray()) + "`n")
}

function Read-OlctlConfig {
    $cfg = @{
        protect_default_folders = $true
        protected_folders       = @('Deleted Items', 'Junk Email', 'Sent Items', 'Outbox')
        max_items               = 200
        default_body_chars      = 20000
        preview_chars           = 240
        from_by_store           = $null
        presets                 = $null
    }
    $path = Join-Path (Get-OlctlHome) 'config.json'
    if (Test-Path -LiteralPath $path) {
        $raw = $null
        try {
            $raw = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
        } catch {
            Invoke-OlFail -Code 'bad_config' -Message ("cannot read " + $path) -Detail $_.Exception.Message
        }
        $parsed = $null
        try {
            $parsed = ConvertFrom-Json -InputObject $raw
        } catch {
            Invoke-OlFail -Code 'bad_config' -Message ($path + ' is not valid JSON') -Detail $_.Exception.Message
        }
        foreach ($p in $parsed.PSObject.Properties) { $cfg[$p.Name] = $p.Value }
    }
    return $cfg
}

function Add-OlctlAudit {
    param([string]$Action, $Detail)
    # Auditing must never break the command it is recording.
    try {
        $dir = Get-OlctlHome
        if (-not (Test-Path -LiteralPath $dir)) {
            [void](New-Item -ItemType Directory -Path $dir -Force)
        }
        $rec = @{
            ts     = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz')
            action = $Action
            detail = $Detail
        }
        $line = ConvertTo-Json -InputObject $rec -Depth 8 -Compress
        $sw = New-Object IO.StreamWriter((Join-Path $dir 'audit.jsonl'), $true, (New-Object Text.UTF8Encoding($false)))
        try { $sw.WriteLine($line) } finally { $sw.Dispose() }
    } catch { }
}

# Scan marks: %USERPROFILE%\.olctl\marks.json, {"<name>": "yyyy-MM-ddTHH:mm:ss"}.
# Local bookkeeping only; nothing in the mailbox changes.
function Get-OlctlMarksPath { return (Join-Path (Get-OlctlHome) 'marks.json') }

function Read-OlctlMarks {
    $marks = @{}
    $path = Get-OlctlMarksPath
    if (-not (Test-Path -LiteralPath $path)) { return $marks }
    try {
        $parsed = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($path, [Text.Encoding]::UTF8))
        foreach ($p in $parsed.PSObject.Properties) { $marks[$p.Name] = [string]$p.Value }
    } catch {
        Invoke-OlFail -Code 'bad_config' -Message ($path + ' is not valid JSON') -Detail $_.Exception.Message
    }
    return $marks
}

function Save-OlctlMark {
    param([string]$Name, [string]$When)
    $marks = Read-OlctlMarks
    $marks[$Name] = $When
    $dir = Get-OlctlHome
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $json = ConvertTo-Json -InputObject $marks -Depth 3
    [IO.File]::WriteAllText((Get-OlctlMarksPath), $json, (New-Object Text.UTF8Encoding($false)))
}

function Test-OlctlMarkName {
    param([string]$Name)
    if (-not ($Name -match '^[A-Za-z0-9_.-]+$')) {
        Invoke-OlFail -Code 'bad_argument' -Message ("mark name '" + $Name + "' may only use letters, digits, '.', '_' and '-'")
    }
}

# ------------------------------------------------------------ time and filters

function ConvertTo-OlctlWhen {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    $t = $Text.Trim()
    $m = [regex]::Match($t, '^(\d+)\s*([mhdwMHDW])$')
    if ($m.Success) {
        $n = [int]$m.Groups[1].Value
        switch ($m.Groups[2].Value.ToLowerInvariant()) {
            'm' { return (Get-Date).AddMinutes(-$n) }
            'h' { return (Get-Date).AddHours(-$n) }
            'd' { return (Get-Date).AddDays(-$n) }
            'w' { return (Get-Date).AddDays(-7 * $n) }
        }
    }
    # [string[]] cast is load-bearing: PowerShell 5.1 passes a raw Object[] to
    # TryParseExact in a way that makes it silently return $false for every
    # absolute date, so --since/--until with YYYY-MM-DD would always fail.
    $formats = [string[]]@('yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd')
    $parsed = [datetime]::MinValue
    $ok = [datetime]::TryParseExact($t, $formats, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None, [ref]$parsed)
    if ($ok) { return $parsed }
    Invoke-OlFail -Code 'bad_argument' -Message ("cannot parse time '" + $Text +
        "'. Use 30m, 6h, 3d, 2w, YYYY-MM-DD or YYYY-MM-DDTHH:MM.")
}

function Get-OlctlRestrictLiteral {
    param([datetime]$When)
    # Outlook's Restrict() wants US-format date literals no matter what the
    # Windows locale is. InvariantCulture is load-bearing here: on a German or
    # Czech machine the local culture would render AM/PM as "vorm."/"dop." and
    # the filter would be rejected.
    return $When.ToString('MM/dd/yyyy hh:mm tt', [Globalization.CultureInfo]::InvariantCulture)
}

# DASL (@SQL=) compares dates in UTC, unlike the Jet syntax above.
function Get-OlctlDaslLiteral {
    param([datetime]$When)
    return $When.ToUniversalTime().ToString('MM/dd/yyyy hh:mm tt', [Globalization.CultureInfo]::InvariantCulture)
}

<#
    The server-side filter. Jet syntax ([UnRead] = True ...) as before; a body
    search needs DASL, and Restrict takes one syntax or the other, never a mix,
    so --body-contains switches the whole filter to @SQL=.
#>
function New-OlctlRestrict {
    param($Opt, $Since, $Until, [string]$BodyContains)
    $clauses = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrEmpty($BodyContains)) {
        if ($Opt.ContainsKey('unread')) { [void]$clauses.Add('[UnRead] = True') }
        if ($Opt.ContainsKey('read')) { [void]$clauses.Add('[UnRead] = False') }
        if ($Opt.ContainsKey('flagged')) { [void]$clauses.Add('[FlagStatus] = 2') }
        if ($null -ne $Since) { [void]$clauses.Add("[ReceivedTime] >= '" + (Get-OlctlRestrictLiteral $Since) + "'") }
        if ($null -ne $Until) { [void]$clauses.Add("[ReceivedTime] <= '" + (Get-OlctlRestrictLiteral $Until) + "'") }
        if ($clauses.Count -eq 0) { return $null }
        return [string]::Join(' AND ', $clauses.ToArray())
    }
    if ($Opt.ContainsKey('unread')) { [void]$clauses.Add('"urn:schemas:httpmail:read" = 0') }
    if ($Opt.ContainsKey('read')) { [void]$clauses.Add('"urn:schemas:httpmail:read" = 1') }
    if ($Opt.ContainsKey('flagged')) { [void]$clauses.Add('"http://schemas.microsoft.com/mapi/proptag/0x10900003" = 2') }
    if ($null -ne $Since) { [void]$clauses.Add('"urn:schemas:httpmail:datereceived" >= ''' + (Get-OlctlDaslLiteral $Since) + '''') }
    if ($null -ne $Until) { [void]$clauses.Add('"urn:schemas:httpmail:datereceived" <= ''' + (Get-OlctlDaslLiteral $Until) + '''') }
    [void]$clauses.Add('"urn:schemas:httpmail:textdescription" LIKE ''%' + $BodyContains.Replace("'", "''") + '%''')
    return ('@SQL=' + [string]::Join(' AND ', $clauses.ToArray()))
}

function ConvertFrom-OlctlCursor {
    param([string]$Text, [bool]$OldestFirst)
    $obj = $null
    try {
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Text))
        $obj = ConvertFrom-Json -InputObject $json
    } catch { $obj = $null }
    if (($null -eq $obj) -or ($null -eq $obj.PSObject.Properties['ref'])) {
        Invoke-OlFail -Code 'bad_argument' -Message '--after-cursor takes the next_cursor value of a previous list, unchanged'
    }
    if ([bool]$obj.o -ne $OldestFirst) {
        Invoke-OlFail -Code 'bad_argument' -Message ('this cursor came from a listing in the other order; ' +
            'repeat --oldest-first exactly as in the first call')
    }
    return @{ r = [string]$obj.r; ref = [string]$obj.ref; t = [string]$obj.t }
}

# Exclusions from --exclude-* plus every --preset named, from config.json:
# "presets": {"noise": {"exclude_sender": [...], "exclude_subject": [...]}}.
function Get-OlctlExclusions {
    param($Opt, $Cfg, [string]$Key)
    $out = New-Object System.Collections.Generic.List[object]
    $direct = Get-OlctlOptList $Opt $Key
    if ($null -ne $direct) { foreach ($v in $direct) { [void]$out.Add([string]$v) } }
    $names = Get-OlctlOptList $Opt 'preset'
    if ($null -ne $names) {
        foreach ($nm in $names) {
            $pr = $null
            if ($null -ne $Cfg['presets']) { $pr = $Cfg['presets'].PSObject.Properties[[string]$nm] }
            if ($null -eq $pr) {
                $known = @()
                if ($null -ne $Cfg['presets']) { $known = @($Cfg['presets'].PSObject.Properties | ForEach-Object { $_.Name }) }
                Invoke-OlFail -Code 'bad_argument' -Extra @{ known_presets = $known } `
                    -Message ("no preset named '" + [string]$nm + "' in " + (Join-Path (Get-OlctlHome) 'config.json'))
            }
            $vals = $pr.Value.PSObject.Properties[$Key]
            if ($null -ne $vals) { foreach ($v in @($vals.Value)) { [void]$out.Add([string]$v) } }
        }
    }
    return , @($out.ToArray())
}

# -------------------------------------------------------------- option parsing

$script:OLCTL_FLAGS = @(
    'unread', 'read', 'flagged', 'has_attachments', 'oldest_first', 'no_preview',
    'any_class', 'count_only', 'headers', 'overwrite', 'all', 'html', 'clear',
    'allow_protected', 'pretty', 'dry_run', 'b64out', 'help', 'h', 'version',
    'mark_read', 'stop', 'disabled', 'replace', 'include_subfolders',
    'quick', 'strip_quoted', 'keep_quoted', 'whole_body', 'stop_on_error'
)
$script:OLCTL_VALUES = @(
    'folder', 'shared', 'since', 'until', 'sender', 'subject', 'category', 'limit',
    'scan_max', 'depth', 'body', 'body_chars', 'out', 'pattern', 'due_code', 'text',
    'to', 'cc', 'attach', 'add', 'remove',
    'name', 'store', 'move_to', 'sender_contains', 'subject_contains', 'from', 'header_contains',
    'json_out', 'file', 'since_mark', 'save_mark', 'after_cursor', 'exclude_sender',
    'exclude_subject', 'preset', 'body_contains', 'save_to', 'supersede', 'also'
)
$script:OLCTL_MULTI = @('to', 'cc', 'attach', 'add', 'remove', 'folder',
    'exclude_sender', 'exclude_subject', 'preset', 'also')

function Get-OlctlParsed {
    param([string[]]$Argv)

    $opt = @{}
    $pos = New-Object System.Collections.Generic.List[object]
    $i = 0
    while ($i -lt $Argv.Count) {
        $tok = [string]$Argv[$i]
        if ($tok -match '^--?[A-Za-z]') {
            $name = ($tok -replace '^--?', '').Replace('-', '_').ToLowerInvariant()
            if ($name -eq 'set') {
                # --set takes zero or more values; zero means "clear them all".
                $vals = New-Object System.Collections.Generic.List[object]
                while (($i + 1) -lt $Argv.Count -and -not ([string]$Argv[$i + 1] -match '^--?[A-Za-z]')) {
                    $i++
                    [void]$vals.Add([string]$Argv[$i])
                }
                $opt['set'] = $vals
            } elseif ($script:OLCTL_FLAGS -contains $name) {
                $opt[$name] = $true
            } elseif ($script:OLCTL_VALUES -contains $name) {
                if (($i + 1) -ge $Argv.Count) {
                    Invoke-OlFail -Code 'bad_argument' -Message ('option ' + $tok + ' needs a value')
                }
                $i++
                $val = [string]$Argv[$i]
                if ($script:OLCTL_MULTI -contains $name) {
                    if (-not $opt.ContainsKey($name)) {
                        $opt[$name] = New-Object System.Collections.Generic.List[object]
                    }
                    [void]$opt[$name].Add($val)
                } else {
                    $opt[$name] = $val
                }
            } else {
                Invoke-OlFail -Code 'bad_argument' -Message ("unknown option '" + $tok +
                    "'. Run olctl --help for the list.")
            }
        } else {
            [void]$pos.Add($tok)
        }
        $i++
    }
    return @{ opt = $opt; pos = @($pos.ToArray()) }
}

# A repeatable option read as a single value means its last occurrence.
function Get-OlctlOptString { param($Opt, [string]$Name, $Default = $null)
    if (-not $Opt.ContainsKey($Name)) { return $Default }
    $v = $Opt[$Name]
    if ($v -is [System.Collections.Generic.List[object]]) { return [string]$v[$v.Count - 1] }
    return [string]$v }

function Get-OlctlOptInt { param($Opt, [string]$Name, [int]$Default)
    if (-not $Opt.ContainsKey($Name)) { return $Default }
    $parsed = 0
    if ([int]::TryParse([string]$Opt[$Name], [ref]$parsed)) { return $parsed }
    Invoke-OlFail -Code 'bad_argument' -Message ('--' + ($Name -replace '_', '-') +
        " needs a whole number, got '" + [string]$Opt[$Name] + "'") }

function Get-OlctlOptList { param($Opt, [string]$Name)
    if (-not $Opt.ContainsKey($Name)) { return $null }
    $v = $Opt[$Name]
    if ($v -is [System.Collections.Generic.List[object]]) { return , @($v.ToArray()) }
    return , @(@($v)) }

function Get-OlctlRequiredRef {
    param($Pos, [string]$Command)
    if ($Pos.Count -lt 2 -or [string]::IsNullOrEmpty([string]$Pos[1])) {
        Invoke-OlFail -Code 'bad_argument' -Message ($Command +
            ' needs a REF, which comes from a previous list or get')
    }
    return [string]$Pos[1]
}

# A literal --text value, a file path, or - for stdin.
function Resolve-OlctlText {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -eq '-') { return [Console]::In.ReadToEnd() }
    $looksLikePath = $false
    try { $looksLikePath = (Test-Path -LiteralPath $Value -PathType Leaf) } catch { $looksLikePath = $false }
    if ($looksLikePath) {
        try { return [IO.File]::ReadAllText($Value, [Text.Encoding]::UTF8) } catch { return $Value }
    }
    return $Value
}

# ------------------------------------------------------------ command mapping

<#
    One command line -> one engine call: @{ op; args; audit; save_mark;
    started }. Pure argument work, no COM, so batch can validate every line
    before anything runs.
#>
function Get-OlctlCall {
    param($Pos, $Opt, $Cfg, [bool]$DryRun, $Policy = $null)

    $cmd = ([string]$Pos[0]).ToLowerInvariant()
    Assert-OlctlPolicy -Cmd $cmd -Policy $Policy -DryRun $DryRun
    $a = @{}
    $op = $null
    $auditAs = $null
    $saveMark = $null
    $started = $null

    switch ($cmd) {
        'doctor' {
            $op = 'doctor'
            $a['quick'] = $Opt.ContainsKey('quick')
        }

        'folders' {
            $op = 'folders'
            $a['folder'] = Get-OlctlOptString $Opt 'folder' $null
            $a['shared'] = Get-OlctlOptString $Opt 'shared' $null
            $a['depth'] = Get-OlctlOptInt $Opt 'depth' 2
        }

        'list' {
            $op = 'list'
            if (-not $Opt.ContainsKey('folder')) {
                Invoke-OlFail -Code 'bad_argument' -Message 'list needs --folder'
            }
            $maxItems = 200
            try { $maxItems = [int]$Cfg['max_items'] } catch { $maxItems = 200 }
            $limit = Get-OlctlOptInt $Opt 'limit' 25
            if ($limit -gt $maxItems) { $limit = $maxItems }
            $previewChars = 240
            try { $previewChars = [int]$Cfg['preview_chars'] } catch { $previewChars = 240 }
            if ($Opt.ContainsKey('no_preview')) { $previewChars = 0 }
            $oldest = $Opt.ContainsKey('oldest_first')

            # All time arithmetic happens before COM is touched, so a bad
            # --since or cursor fails fast.
            $since = $null
            if ($Opt.ContainsKey('since_mark')) {
                $markName = Get-OlctlOptString $Opt 'since_mark' $null
                Test-OlctlMarkName $markName
                $marks = Read-OlctlMarks
                if ($marks.ContainsKey($markName)) {
                    $since = ConvertTo-OlctlWhen $marks[$markName]
                } elseif ($Opt.ContainsKey('since')) {
                    $since = ConvertTo-OlctlWhen (Get-OlctlOptString $Opt 'since' $null)
                } else {
                    Invoke-OlFail -Code 'bad_argument' -Extra @{ known_marks = @($marks.Keys) } `
                        -Message ("no mark named '" + $markName + "' yet. Pass --since as well for the first run; it is used until the mark exists.")
                }
            } elseif ($Opt.ContainsKey('since')) {
                $since = ConvertTo-OlctlWhen (Get-OlctlOptString $Opt 'since' $null)
            }
            $until = $null
            if ($Opt.ContainsKey('until')) { $until = ConvertTo-OlctlWhen (Get-OlctlOptString $Opt 'until' $null) }

            $started = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
            $cursor = $null
            if ($Opt.ContainsKey('after_cursor')) {
                $cursor = ConvertFrom-OlctlCursor (Get-OlctlOptString $Opt 'after_cursor' '') $oldest
                if (-not [string]::IsNullOrEmpty($cursor['t'])) { $started = $cursor['t'] }
                # Narrow the server filter to the cursor too. Restrict compares
                # whole minutes, hence the one-minute margin; the exact cut is
                # made per item in the engine.
                if (-not [string]::IsNullOrEmpty($cursor['r'])) {
                    $cr = ConvertTo-OlctlWhen $cursor['r']
                    if ($oldest) {
                        $b = $cr.AddMinutes(-1)
                        if (($null -eq $since) -or ($b -gt $since)) { $since = $b }
                    } else {
                        $b = $cr.AddMinutes(1)
                        if (($null -eq $until) -or ($b -lt $until)) { $until = $b }
                    }
                }
                $a['cursor'] = @{ r = $cursor['r']; ref = $cursor['ref'] }
            }
            if ($Opt.ContainsKey('save_mark')) {
                $saveMark = Get-OlctlOptString $Opt 'save_mark' $null
                Test-OlctlMarkName $saveMark
            }

            $a['folders'] = Get-OlctlOptList $Opt 'folder'
            $a['shared'] = Get-OlctlOptString $Opt 'shared' $null
            $a['restrict'] = New-OlctlRestrict -Opt $Opt -Since $since -Until $until `
                -BodyContains (Get-OlctlOptString $Opt 'body_contains' $null)
            $a['limit'] = $limit
            $a['scan_max'] = Get-OlctlOptInt $Opt 'scan_max' 2000
            $a['oldest_first'] = $oldest
            $a['preview_chars'] = $previewChars
            $a['any_class'] = $Opt.ContainsKey('any_class')
            $a['count_only'] = $Opt.ContainsKey('count_only')
            $a['include_subfolders'] = $Opt.ContainsKey('include_subfolders')
            $a['started'] = $started
            $a['filters'] = @{
                sender          = Get-OlctlOptString $Opt 'sender' $null
                subject         = Get-OlctlOptString $Opt 'subject' $null
                category        = Get-OlctlOptString $Opt 'category' $null
                has_attachments = $Opt.ContainsKey('has_attachments')
                exclude_sender  = Get-OlctlExclusions $Opt $Cfg 'exclude_sender'
                exclude_subject = Get-OlctlExclusions $Opt $Cfg 'exclude_subject'
            }
        }

        'get' {
            $op = 'get'
            $bodyChars = 20000
            try { $bodyChars = [int]$Cfg['default_body_chars'] } catch { $bodyChars = 20000 }
            $a['ref'] = Get-OlctlRequiredRef $Pos 'get'
            $a['body'] = Get-OlctlOptString $Opt 'body' 'text'
            if (@('text', 'html', 'none') -notcontains $a['body']) {
                Invoke-OlFail -Code 'bad_argument' -Message '--body must be text, html or none'
            }
            if ($Opt.ContainsKey('strip_quoted') -and ($a['body'] -ne 'text')) {
                Invoke-OlFail -Code 'bad_argument' -Message '--strip-quoted only works with --body text'
            }
            $a['body_chars'] = Get-OlctlOptInt $Opt 'body_chars' $bodyChars
            $a['headers'] = $Opt.ContainsKey('headers')
            $a['strip_quoted'] = $Opt.ContainsKey('strip_quoted')
        }

        'thread' {
            $op = 'thread'
            $a['ref'] = Get-OlctlRequiredRef $Pos 'thread'
            $a['body'] = Get-OlctlOptString $Opt 'body' 'text'
            if (@('text', 'none') -notcontains $a['body']) {
                Invoke-OlFail -Code 'bad_argument' -Message 'thread --body must be text or none'
            }
            $a['body_chars'] = Get-OlctlOptInt $Opt 'body_chars' 4000
            $a['strip_quoted'] = -not $Opt.ContainsKey('keep_quoted')
            $a['also'] = Get-OlctlOptList $Opt 'also'
            if ($Opt.ContainsKey('since')) {
                $a['restrict'] = New-OlctlRestrict -Opt @{} -Since (ConvertTo-OlctlWhen (Get-OlctlOptString $Opt 'since' $null)) -Until $null
            }
        }

        'save-attachments' {
            $op = 'save_attachments'
            if (-not $Opt.ContainsKey('out')) {
                Invoke-OlFail -Code 'bad_argument' -Message 'save-attachments needs --out'
            }
            $a['ref'] = Get-OlctlRequiredRef $Pos 'save-attachments'
            $a['out'] = Get-OlctlOptString $Opt 'out' $null
            $a['pattern'] = Get-OlctlOptString $Opt 'pattern' $null
            $a['overwrite'] = $Opt.ContainsKey('overwrite')
            $a['dry_run'] = $DryRun
            $auditAs = 'save-attachments'
        }

        'move' {
            $op = 'move'
            if (-not $Opt.ContainsKey('to')) {
                Invoke-OlFail -Code 'bad_argument' -Message 'move needs --to'
            }
            $toList = Get-OlctlOptList $Opt 'to'
            $a['ref'] = Get-OlctlRequiredRef $Pos 'move'
            $a['to'] = [string]$toList[0]
            $a['shared'] = Get-OlctlOptString $Opt 'shared' $null
            $a['allow_protected'] = $Opt.ContainsKey('allow_protected')
            $a['protect_defaults'] = [bool]$Cfg['protect_default_folders']
            $a['protected_names'] = @($Cfg['protected_folders'])
            $a['dry_run'] = $DryRun
            $auditAs = 'move'
        }

        'mark' {
            $op = 'mark'
            $a['ref'] = Get-OlctlRequiredRef $Pos 'mark'
            if ($Pos.Count -lt 3) {
                Invoke-OlFail -Code 'bad_argument' -Message 'mark needs read or unread'
            }
            $state = ([string]$Pos[2]).ToLowerInvariant()
            if (@('read', 'unread') -notcontains $state) {
                Invoke-OlFail -Code 'bad_argument' -Message "mark takes 'read' or 'unread'"
            }
            $a['unread'] = ($state -eq 'unread')
            $a['dry_run'] = $DryRun
            $auditAs = 'mark'
        }

        'flag' {
            $op = 'flag'
            $a['ref'] = Get-OlctlRequiredRef $Pos 'flag'
            $a['clear'] = $Opt.ContainsKey('clear')
            $a['text'] = Get-OlctlOptString $Opt 'text' 'Follow up'
            $a['due_code'] = Get-OlctlOptInt $Opt 'due_code' 3
            $a['dry_run'] = $DryRun
            $auditAs = 'flag'
        }

        'categorize' {
            $op = 'categorize'
            $a['ref'] = Get-OlctlRequiredRef $Pos 'categorize'
            $a['add'] = Get-OlctlOptList $Opt 'add'
            $a['remove'] = Get-OlctlOptList $Opt 'remove'
            if ($Opt.ContainsKey('set')) {
                $setList = $Opt['set']
                $a['set'] = @($setList.ToArray())
            }
            $a['dry_run'] = $DryRun
            $auditAs = 'categorize'
        }

        'draft-reply' {
            $op = 'draft_reply'
            if (-not $Opt.ContainsKey('text')) {
                Invoke-OlFail -Code 'bad_argument' -Message 'draft-reply needs --text'
            }
            $a['ref'] = Get-OlctlRequiredRef $Pos 'draft-reply'
            $a['text'] = Resolve-OlctlText (Get-OlctlOptString $Opt 'text' '')
            $a['reply_all'] = $Opt.ContainsKey('all')
            $a['html'] = $Opt.ContainsKey('html')
            $a['subject'] = Get-OlctlOptString $Opt 'subject' $null
            $a['attach'] = Get-OlctlOptList $Opt 'attach'
            $a['from'] = Get-OlctlOptString $Opt 'from' $null
            $a['save_to'] = Get-OlctlOptString $Opt 'save_to' $null
            $a['supersede'] = Get-OlctlOptString $Opt 'supersede' $null
            $a['from_by_store'] = $Cfg['from_by_store']
            $a['dry_run'] = $DryRun
            $auditAs = 'draft-reply'
        }

        'draft-new' {
            $op = 'draft_new'
            if (-not $Opt.ContainsKey('to')) {
                Invoke-OlFail -Code 'bad_argument' -Message 'draft-new needs at least one --to'
            }
            if (-not $Opt.ContainsKey('text')) {
                Invoke-OlFail -Code 'bad_argument' -Message 'draft-new needs --text'
            }
            $a['to'] = Get-OlctlOptList $Opt 'to'
            $a['cc'] = Get-OlctlOptList $Opt 'cc'
            $a['subject'] = Get-OlctlOptString $Opt 'subject' $null
            $a['text'] = Resolve-OlctlText (Get-OlctlOptString $Opt 'text' '')
            $a['html'] = $Opt.ContainsKey('html')
            $a['attach'] = Get-OlctlOptList $Opt 'attach'
            $a['from'] = Get-OlctlOptString $Opt 'from' $null
            $a['save_to'] = Get-OlctlOptString $Opt 'save_to' $null
            $a['dry_run'] = $DryRun
            $auditAs = 'draft-new'
        }

        'draft-update' {
            $op = 'draft_update'
            $a['ref'] = Get-OlctlRequiredRef $Pos 'draft-update'
            if (-not ($Opt.ContainsKey('text') -or $Opt.ContainsKey('subject') -or
                    $Opt.ContainsKey('attach') -or $Opt.ContainsKey('from'))) {
                Invoke-OlFail -Code 'bad_argument' -Message 'draft-update needs --text, --subject, --from or --attach'
            }
            if ($Opt.ContainsKey('whole_body') -and (-not $Opt.ContainsKey('text'))) {
                Invoke-OlFail -Code 'bad_argument' -Message '--whole-body needs --text'
            }
            if ($Opt.ContainsKey('text')) { $a['text'] = Resolve-OlctlText (Get-OlctlOptString $Opt 'text' '') }
            $a['html'] = $Opt.ContainsKey('html')
            $a['whole_body'] = $Opt.ContainsKey('whole_body')
            $a['subject'] = Get-OlctlOptString $Opt 'subject' $null
            $a['from'] = Get-OlctlOptString $Opt 'from' $null
            $a['attach'] = Get-OlctlOptList $Opt 'attach'
            $a['dry_run'] = $DryRun
            $auditAs = 'draft-update'
        }

        'folder-add' {
            $op = 'folder_add'
            if (-not $Opt.ContainsKey('folder')) {
                Invoke-OlFail -Code 'bad_argument' -Message 'folder-add needs --folder (the parent path)'
            }
            if (-not $Opt.ContainsKey('name')) {
                Invoke-OlFail -Code 'bad_argument' -Message 'folder-add needs --name'
            }
            $chain = @()
            foreach ($v in (Get-OlctlOptList $Opt 'name')) {
                foreach ($seg in ([string]$v -split '/')) {
                    $seg2 = ([string]$seg).Trim()
                    if ($seg2.Length -gt 0) { $chain += $seg2 }
                }
            }
            if ($chain.Count -eq 0) {
                Invoke-OlFail -Code 'bad_argument' -Message 'folder-add needs a non-empty --name'
            }
            $a['folder'] = Get-OlctlOptString $Opt 'folder' $null
            $a['name'] = $chain
            $a['shared'] = Get-OlctlOptString $Opt 'shared' $null
            $a['dry_run'] = $DryRun
            $auditAs = 'folder-add'
        }

        'rules-list' {
            $op = 'rules_list'
            $a['store'] = Get-OlctlOptString $Opt 'store' $null
        }

        'rules-add' {
            $op = 'rules_add'
            $a['name'] = Get-OlctlOptString $Opt 'name' $null
            $a['store'] = Get-OlctlOptString $Opt 'store' $null
            $senders = @()
            foreach ($v in (Get-OlctlOptList $Opt 'sender_contains')) { $senders += @(([string]$v -split ',')) }
            $subjects = @()
            foreach ($v in (Get-OlctlOptList $Opt 'subject_contains')) { $subjects += @(([string]$v -split ',')) }
            $headers = @()
            foreach ($v in (Get-OlctlOptList $Opt 'header_contains')) { $headers += @(([string]$v -split ',')) }
            $froms = @()
            foreach ($v in (Get-OlctlOptList $Opt 'from')) { $froms += @(([string]$v -split ',')) }
            $a['sender_contains'] = @(($senders | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_.Length -gt 0 }))
            $a['subject_contains'] = @(($subjects | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_.Length -gt 0 }))
            $a['header_contains'] = @(($headers | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_.Length -gt 0 }))
            $a['from'] = @(($froms | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_.Length -gt 0 }))
            $a['move_to'] = Get-OlctlOptString $Opt 'move_to' $null
            $a['shared'] = Get-OlctlOptString $Opt 'shared' $null
            $a['mark_read'] = $Opt.ContainsKey('mark_read')
            $a['stop'] = $Opt.ContainsKey('stop')
            $a['enabled'] = -not $Opt.ContainsKey('disabled')
            $a['replace'] = $Opt.ContainsKey('replace')
            $a['dry_run'] = $DryRun
            $auditAs = 'rules-add'
        }

        'rules-delete' {
            $op = 'rules_delete'
            $a['name'] = Get-OlctlOptString $Opt 'name' $null
            $a['store'] = Get-OlctlOptString $Opt 'store' $null
            $a['dry_run'] = $DryRun
            $auditAs = 'rules-delete'
        }

        'rules-run' {
            $op = 'rules_run'
            $a['name'] = Get-OlctlOptString $Opt 'name' $null
            $a['all'] = $Opt.ContainsKey('all')
            if ([string]::IsNullOrEmpty($a['name']) -and (-not $a['all'])) {
                Invoke-OlFail -Code 'bad_argument' -Message 'rules-run needs --name or --all'
            }
            $a['store'] = Get-OlctlOptString $Opt 'store' $null
            $a['folder'] = Get-OlctlOptString $Opt 'folder' $null
            $a['shared'] = Get-OlctlOptString $Opt 'shared' $null
            $a['include_subfolders'] = $Opt.ContainsKey('include_subfolders')
            $a['dry_run'] = $DryRun
            $auditAs = 'rules-run'
        }

        default {
            Invoke-OlFail -Code 'bad_argument' -Message ("unknown command '" + $cmd +
                "'. Run olctl --help for the list.")
        }
    }

    # The engine reads its arguments with PSObject.Properties, so hand it a
    # PSCustomObject. Round-tripping through JSON converts the nested
    # hashtables and lists in one step.
    $opArgs = $null
    if ($a.Count -gt 0) {
        $opArgs = ConvertFrom-Json -InputObject (ConvertTo-Json -InputObject $a -Depth 10)
    }
    return @{ op = $op; args = $opArgs; audit = $auditAs; save_mark = $saveMark; started = $started; dry_run = $DryRun }
}

# Run one mapped call: engine, audit log, scan mark.
function Invoke-OlctlRun {
    param($Call, $Conn = $null)
    $captured = @(Invoke-OlOp -Op $Call['op'] -OpArgs $Call['args'] -Conn $Conn)
    $data = $null
    for ($k = $captured.Count - 1; $k -ge 0; $k--) {
        if ($captured[$k] -is [hashtable]) { $data = $captured[$k]; break }
    }
    if ($null -eq $data) { $data = @{} }

    if (($null -ne $Call['audit']) -and (-not $Call['dry_run'])) {
        Add-OlctlAudit -Action $Call['audit'] -Detail $data
    }

    if (-not [string]::IsNullOrEmpty($Call['save_mark'])) {
        # Only a complete scan may move the mark; a truncated one would make
        # the next --since-mark skip what this run never reached.
        if ($data['truncated'] -eq $true) {
            $data['mark_saved'] = $null
            $data['mark_note'] = ('not saved: the result is truncated. Page on with --after-cursor ' +
                '(and the same --save-mark); the mark is saved by the last page.')
        } else {
            Save-OlctlMark -Name $Call['save_mark'] -When $Call['started']
            $data['mark_saved'] = @{ name = $Call['save_mark']; at = $Call['started'] }
        }
    }
    return $data
}

# The error envelope for a caught failure: the structured payload when there
# is one, otherwise the bare exception.
function Get-OlctlErrorEnvelope {
    param($Err)
    $payload = $null
    $target = $null
    try { $target = $Err.TargetObject } catch { $target = $null }
    if (($target -is [hashtable]) -and $target.ContainsKey('__olctl_error')) {
        $payload = $target
    } elseif (($null -ne $script:OlErrorPayload) -and ($script:OlErrorPayload -is [hashtable])) {
        $payload = $script:OlErrorPayload
    }

    $envelope = @{}
    $envelope['ok'] = $false
    if ($null -ne $payload) {
        foreach ($key in @($payload.Keys)) {
            if ($key -eq '__olctl_error') { continue }
            $envelope[$key] = $payload[$key]
        }
        if (-not $envelope.ContainsKey('code')) { $envelope['code'] = 'error' }
        if (-not $envelope.ContainsKey('error')) { $envelope['error'] = 'unspecified error' }
        if (-not $envelope.ContainsKey('detail')) { $envelope['detail'] = $null }
    } else {
        $msg = 'unexpected failure'
        try { $msg = [string]$Err.Exception.Message } catch { }
        if ([string]::IsNullOrEmpty($msg)) { $msg = 'unexpected failure' }
        $detail = $null
        try {
            $detail = ($Err.Exception.GetType().FullName + ' at line ' +
                [string]$Err.InvocationInfo.ScriptLineNumber)
        } catch { $detail = $null }
        $envelope['code'] = 'error'
        $envelope['error'] = $msg
        $envelope['detail'] = $detail
    }
    return $envelope
}

<#
    batch: one JSON array of argv strings per line, e.g.
        ["mark", "<ref>", "read"]
        ["move", "<ref>", "--to", "projects/Inbox/Processed"]
    Every line is validated first; if any is malformed nothing runs. Then all
    run on one Outlook attach. A global --dry-run applies to every line.
#>
function Invoke-OlctlBatch {
    param($Opt, $Cfg, [bool]$DryRun, $Policy = $null)

    $file = Get-OlctlOptString $Opt 'file' '-'
    $raw = ''
    if ($file -eq '-') {
        $raw = [Console]::In.ReadToEnd()
    } else {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
            Invoke-OlFail -Code 'bad_argument' -Message ("batch file '" + $file + "' not found")
        }
        $raw = [IO.File]::ReadAllText($file, [Text.Encoding]::UTF8)
    }

    $calls = New-Object System.Collections.Generic.List[object]
    $problems = New-Object System.Collections.Generic.List[object]
    $lineNo = 0
    foreach ($ln in [regex]::Split([string]$raw, '\r?\n')) {
        $lineNo = $lineNo + 1
        $t = $ln.Trim()
        if (($t.Length -eq 0) -or $t.StartsWith('#')) { continue }
        $script:OlErrorPayload = $null
        try {
            $argvObj = $null
            try { $argvObj = ConvertFrom-Json -InputObject $t } catch { $argvObj = $null }
            if (-not ($argvObj -is [System.Array]) -or ($argvObj.Count -eq 0)) {
                Invoke-OlFail -Code 'bad_argument' -Message 'each line must be a JSON array of strings, e.g. ["mark", "<ref>", "read"]'
            }
            $lineArgv = [string[]]@($argvObj | ForEach-Object { [string]$_ })
            if (@('batch') -contains $lineArgv[0].ToLowerInvariant()) {
                Invoke-OlFail -Code 'bad_argument' -Message 'batch cannot be nested'
            }
            $p = Get-OlctlParsed -Argv $lineArgv
            $lineDry = $DryRun -or $p['opt'].ContainsKey('dry_run')
            $call = Get-OlctlCall -Pos $p['pos'] -Opt $p['opt'] -Cfg $Cfg -DryRun $lineDry -Policy $Policy
            [void]$calls.Add(@{ line = $lineNo; argv = $lineArgv; call = $call })
        } catch {
            $env = Get-OlctlErrorEnvelope $_
            [void]$problems.Add(@{ line = $lineNo; code = $env['code']; error = $env['error'] })
        }
    }
    if ($problems.Count -gt 0) {
        Invoke-OlFail -Code 'bad_argument' -Extra @{ lines = @($problems.ToArray()) } `
            -Message ([string]$problems.Count + ' batch line(s) are malformed; nothing was run')
    }
    if ($calls.Count -eq 0) {
        Invoke-OlFail -Code 'bad_argument' -Message 'the batch has no commands'
    }

    $conn = Connect-Outlook
    $stopOnError = $Opt.ContainsKey('stop_on_error')
    $results = New-Object System.Collections.Generic.List[object]
    $okCount = 0
    $failCount = 0
    foreach ($c in $calls) {
        $script:OlErrorPayload = $null
        $res = @{ line = $c['line']; command = $c['argv'][0] }
        try {
            $res['data'] = Invoke-OlctlRun -Call $c['call'] -Conn $conn
            $res['ok'] = $true
            $okCount = $okCount + 1
        } catch {
            $env = Get-OlctlErrorEnvelope $_
            $res['ok'] = $false
            $res['error'] = $env
            $failCount = $failCount + 1
            if ($stopOnError) {
                [void]$results.Add($res)
                break
            }
        }
        [void]$results.Add($res)
    }

    $data = @{}
    $data['count'] = $calls.Count
    $data['succeeded'] = $okCount
    $data['failed'] = $failCount
    $data['not_run'] = $calls.Count - $results.Count
    $data['dry_run'] = [bool]$DryRun
    $data['results'] = @($results.ToArray())
    return $data
}

# ---------------------------------------------------------------------- main

$script:Envelope = $null
$script:Pretty = $false
$script:B64Out = $false
$script:JsonOut = $null

try {
    $argv = @()
    if ($null -ne $CliArgs) { $argv = @($CliArgs) }

    $parsed = Get-OlctlParsed -Argv $argv
    $opt = $parsed['opt']
    $pos = $parsed['pos']
    $script:Pretty = $opt.ContainsKey('pretty')
    $script:B64Out = $opt.ContainsKey('b64out')
    $script:JsonOut = Get-OlctlOptString $opt 'json_out' $null
    $dryRun = $opt.ContainsKey('dry_run')

    if ($opt.ContainsKey('version')) {
        $script:Envelope = @{ ok = $true; data = @{ olctl_version = $script:OLCTL_VERSION } }
    } elseif ($opt.ContainsKey('help') -or $opt.ContainsKey('h') -or $pos.Count -eq 0) {
        # Help is for humans, so it is plain text rather than JSON.
        Write-Output $script:USAGE
        exit 0
    } else {
        $policy = Read-OlctlPolicy
        $first = ([string]$pos[0]).ToLowerInvariant()
        if ($first -eq 'guide') {
            # Plain text like --help: it is read, not parsed.
            $guideText = Get-OlctlGuide -Policy $policy
            if ($script:B64Out) {
                Write-Output ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($guideText)))
            } else {
                try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
                Write-Output $guideText
            }
            exit 0
        }
        $cfg = Read-OlctlConfig
        $data = $null
        if ($first -eq 'batch') {
            $data = Invoke-OlctlBatch -Opt $opt -Cfg $cfg -DryRun $dryRun -Policy $policy
        } else {
            $call = Get-OlctlCall -Pos $pos -Opt $opt -Cfg $cfg -DryRun $dryRun -Policy $policy
            $data = Invoke-OlctlRun -Call $call
            if ($first -eq 'doctor') {
                $data['policy'] = @{ read_only = [bool]$policy['read_only'];
                    disabled_commands = @($policy['disabled_commands']); sources = @($policy['sources']) }
            }
        }

        $envelope = @{}
        $envelope['ok'] = $true
        $envelope['data'] = $data
        $script:Envelope = $envelope
    }
} catch {
    $script:Envelope = Get-OlctlErrorEnvelope $_
}

if ($null -eq $script:Envelope) {
    $script:Envelope = @{ ok = $false; code = 'error'; error = 'no response was produced'; detail = $null }
}

# Successful runs print the payload itself, so `olctl list ... | jq .items` works
# without unwrapping. Failures print the error envelope and exit 1.
$out = $script:Envelope
$exit = 0
if ($script:Envelope['ok'] -eq $true) {
    $out = $script:Envelope['data']
} else {
    $exit = 1
}

try {
    if ($script:Pretty) {
        $outJson = ConvertTo-Json -InputObject $out -Depth 12
    } else {
        $outJson = ConvertTo-Json -InputObject $out -Depth 12 -Compress
    }
    # --json-out: the payload goes to the file as UTF-8 without BOM (Windows
    # PowerShell's > would write UTF-16), and stdout gets only where it went.
    # Errors still print in full and leave the file alone.
    if ((-not [string]::IsNullOrEmpty($script:JsonOut)) -and ($exit -eq 0)) {
        try {
            $full = [System.IO.Path]::GetFullPath($script:JsonOut)
            $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($outJson)
            [IO.File]::WriteAllBytes($full, $bytes)
            $outJson = ConvertTo-Json -InputObject @{ json_out = $full; bytes = $bytes.Length } -Compress
        } catch {
            $exit = 1
            $outJson = ConvertTo-Json -Compress -InputObject @{ ok = $false; code = 'no_access';
                error = ("could not write '" + [string]$script:JsonOut + "'"); detail = $_.Exception.Message }
        }
    }
    # Write-Output, not [Console]::Out.Write. The latter writes straight to the
    # console handle, which bypasses the PowerShell pipeline entirely -- so when
    # the script is dot-invoked in-process (.\olctl.ps1 ... | ConvertFrom-Json)
    # the caller captures nothing and the JSON just lands on screen. Going
    # through the pipeline works both in-process and as a child process.
    if ($script:B64Out) {
        # Used by the WSL wrapper: base64 removes every console-codepage question
        # about Czech and German characters on the way out.
        Write-Output ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($outJson)))
    } else {
        try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
        Write-Output $outJson
    }
    exit $exit
} catch {
    try {
        $fallback = '{"ok":false,"code":"error","error":"the response could not be serialised to JSON","detail":null}'
        if ($script:B64Out) {
            Write-Output ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($fallback)))
        } else {
            Write-Output $fallback
        }
        exit 1
    } catch { exit 3 }
}
