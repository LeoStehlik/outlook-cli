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

# Folder types that must never be a move destination without an override.
$script:OL_PROTECTED_TYPES = @(3, 4, 5, 23)

$script:PR_SMTP_ADDRESS = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
$script:PR_INTERNET_HEADERS = 'http://schemas.microsoft.com/mapi/proptag/0x007D001E'
$script:PR_INTERNET_MESSAGE_ID = 'http://schemas.microsoft.com/mapi/proptag/0x1035001E'

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
        foreach ($candidate in @('Archive', 'Archiv', 'Archiv1', 'Archív')) {
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

        $entry['default_folders'] = Get-OlStoreDefaultsMap $storeObj
        $rootChildren = Get-OlChildNames $root
        $entry['top_level_folders'] = @($rootChildren)

        [void]$stores.Add($entry)
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

function Invoke-OlList {
    param($Ol, $Ns, $OpArgs)

    $folderArg = Get-OlString $OpArgs 'folder' $null
    $sharedArg = Get-OlString $OpArgs 'shared' $null
    $restrict = Get-OlString $OpArgs 'restrict' $null
    $limit = Get-OlInt $OpArgs 'limit' 25
    $scanMax = Get-OlInt $OpArgs 'scan_max' 2000
    $oldestFirst = Get-OlBool $OpArgs 'oldest_first' $false
    $previewChars = Get-OlInt $OpArgs 'preview_chars' 0
    $anyClass = Get-OlBool $OpArgs 'any_class' $false
    $countOnly = Get-OlBool $OpArgs 'count_only' $false
    $filters = Get-OlArg $OpArgs 'filters' $null

    $path = $folderArg
    if ($null -eq $path) { $path = '' }
    $folder = Resolve-OlFolder -Ns $Ns -Path $path -Shared $sharedArg
    $folderPath = Get-OlFolderPath -Folder $folder

    $items = $null
    try {
        $items = $folder.Items
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message `
            -Message ("could not open the items of '" + [string]$folderPath + "'")
    }

    $coll = $items
    if (-not [string]::IsNullOrEmpty($restrict)) {
        try {
            $coll = $items.Restrict($restrict)
        } catch {
            Invoke-OlFail -Code 'bad_filter' -Detail $_.Exception.Message `
                -Message ("Outlook rejected the filter '" + $restrict + "'")
        }
    }

    try { $coll.Sort('[ReceivedTime]', (-not $oldestFirst)) } catch { }

    $filterOut = $restrict
    if ([string]::IsNullOrEmpty($filterOut)) { $filterOut = $null }

    if ($countOnly) {
        $c = $null
        try {
            $c = [int]$coll.Count
        } catch {
            Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not count items'
        }
        $data = @{}
        $data['folder'] = $folderPath
        $data['filter'] = $filterOut
        $data['count'] = $c
        return $data
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $scanned = 0
    $skipped = 0

    $it = $null
    try { $it = $coll.GetFirst() } catch { $it = $null }

    while (($null -ne $it) -and ($rows.Count -lt $limit) -and ($scanned -lt $scanMax)) {
        $scanned = $scanned + 1
        try {
            $cls = 0
            try { $cls = [int](Get-Safe { $it.Class } 0) } catch { $cls = 0 }
            if (($cls -ne $script:OL_MAIL_ITEM) -and (-not $anyClass)) {
                $skipped = $skipped + 1
            } else {
                $row = Get-OlRow -Item $it -PreviewChars $previewChars
                if (Test-OlRowFilters -Row $row -Filters $filters) {
                    [void]$rows.Add($row)
                }
            }
        } catch {
            $skipped = $skipped + 1
        }
        try { $it = $coll.GetNext() } catch { $it = $null }
    }

    $truncated = (($scanned -ge $scanMax) -or ($rows.Count -ge $limit))

    $data = @{}
    $data['folder'] = $folderPath
    $data['filter'] = $filterOut
    $data['returned'] = $rows.Count
    $data['scanned'] = $scanned
    $data['skipped_non_mail_or_unreadable'] = $skipped
    $data['truncated'] = [bool]$truncated
    $data['items'] = @($rows.ToArray())
    return $data
}

# ----------------------------------------------------------------------- get

function Invoke-OlGet {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $bodyMode = Get-OlString $OpArgs 'body' 'text'
    $bodyChars = Get-OlInt $OpArgs 'body_chars' 20000
    $wantHeaders = Get-OlBool $OpArgs 'headers' $false

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
    $row['internet_message_id'] = Get-Safe { $item.PropertyAccessor.GetProperty($script:PR_INTERNET_MESSAGE_ID) } $null

    return $row
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

# --------------------------------------------------------------- draft_reply

function Invoke-OlDraftReply {
    param($Ol, $Ns, $OpArgs)

    $ref = Get-OlString $OpArgs 'ref' $null
    $text = Get-OlString $OpArgs 'text' ''
    if ($null -eq $text) { $text = '' }
    $replyAll = Get-OlBool $OpArgs 'reply_all' $false
    $asHtml = Get-OlBool $OpArgs 'html' $false
    $subjectOverride = Get-OlString $OpArgs 'subject' $null
    $dryRun = Get-OlBool $OpArgs 'dry_run' $false

    $item = Get-OlItem -Ns $Ns -Ref $ref

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['ref'] = $ref
        $data['reply_all'] = [bool]$replyAll
        $data['chars'] = $text.Length
        $data['subject'] = Get-Safe { $item.Subject } $null
        $data['saved_to'] = 'Drafts'
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

    try {
        if ($asHtml) {
            $existing = [string](Get-Safe { $reply.HTMLBody } '')
            $reply.HTMLBody = $text + $existing
        } else {
            $existing = [string](Get-Safe { $reply.Body } '')
            $reply.Body = $text + "`n" + $existing
        }
        if (-not [string]::IsNullOrEmpty($subjectOverride)) { $reply.Subject = $subjectOverride }
        # Lands in Drafts. This tool never sends: there is no Send() call here.
        $reply.Save()
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not save the reply draft'
    }

    $data = @{}
    $data['draft_ref'] = Get-Safe { Get-OlRef -Item $reply } $null
    $data['in_reply_to'] = $ref
    $data['subject'] = Get-Safe { $reply.Subject } $null
    $data['saved_to'] = 'Drafts'
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

    if ($dryRun) {
        $data = @{}
        $data['dry_run'] = $true
        $data['to'] = @($to)
        $data['cc'] = @($cc)
        $data['subject'] = $subject
        $data['chars'] = $text.Length
        $data['attach'] = @($attach)
        $data['saved_to'] = 'Drafts'
        $data['sent'] = $false
        return $data
    }

    $mail = $null
    try {
        $mail = $Ol.CreateItem(0)
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not create a new mail item'
    }

    try {
        $mail.To = [string]::Join('; ', @($to))
        if ($cc.Count -gt 0) { $mail.CC = [string]::Join('; ', @($cc)) }
        $mail.Subject = $subject
        if ($asHtml) { $mail.HTMLBody = $text } else { $mail.Body = $text }
    } catch {
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

    try {
        # Lands in Drafts. This tool never sends: there is no Send() call here.
        $mail.Save()
    } catch {
        Invoke-OlFail -Code 'error' -Detail $_.Exception.Message -Message 'could not save the new draft'
    }

    $data = @{}
    $data['draft_ref'] = Get-Safe { Get-OlRef -Item $mail } $null
    $data['subject'] = Get-Safe { $mail.Subject } $subject
    $data['saved_to'] = 'Drafts'
    $data['sent'] = $false
    $data['dry_run'] = $false
    return $data
}

# ------------------------------------------------------------------ dispatch

$script:OL_OPS = @('doctor', 'folders', 'list', 'get', 'save_attachments', 'move',
    'mark', 'flag', 'categorize', 'draft_reply', 'draft_new')

function Invoke-OlOp {
    param([string]$Op, $OpArgs)

    # Validated before touching COM, so an unknown op always reports
    # bad_argument rather than whatever the COM attach happens to say.
    if (-not ($script:OL_OPS -contains $Op)) {
        Invoke-OlFail -Code 'bad_argument' -Message ("unknown op '" + [string]$Op + "'. Known ops: " +
            [string]::Join(', ', $script:OL_OPS) + '.')
    }

    $conn = Connect-Outlook
    $ol = $conn['app']
    $ns = $conn['ns']

    $data = $null
    switch ($Op) {
        'doctor' { $data = Invoke-OlDoctor -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'folders' { $data = Invoke-OlFolders -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'list' { $data = Invoke-OlList -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'get' { $data = Invoke-OlGet -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'save_attachments' { $data = Invoke-OlSaveAttachments -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'move' { $data = Invoke-OlMove -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'mark' { $data = Invoke-OlMark -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'flag' { $data = Invoke-OlFlag -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'categorize' { $data = Invoke-OlCategorize -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'draft_reply' { $data = Invoke-OlDraftReply -Ol $ol -Ns $ns -OpArgs $OpArgs }
        'draft_new' { $data = Invoke-OlDraftNew -Ol $ol -Ns $ns -OpArgs $OpArgs }
        default {
            Invoke-OlFail -Code 'bad_argument' -Message ("unknown op '" + [string]$Op + "'. Known ops: doctor, folders, list, get, " +
                'save_attachments, move, mark, flag, categorize, draft_reply, draft_new.')
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

$script:OLCTL_VERSION = '0.3.3'

$script:USAGE = @'
olctl <command> [options]     local Outlook Classic CLI -- one JSON object per run

Commands
  doctor                                  connection, stores, protected folders
  folders [--folder P] [--depth N]         folder tree with counts
  list --folder P [filters]                messages, newest first
  get REF [--body text|html|none]          one message in full
  save-attachments REF --out DIR           write attachments to disk
  move REF --to P                          move between folders
  mark REF read|unread                     read state
  flag REF [--text T] [--clear]            follow-up flag
  categorize REF [--add C] [--remove C] [--set A B ...]
  draft-reply REF --text ...               reply saved to Drafts, never sent
  draft-new --to X --subject S --text ...  new mail saved to Drafts, never sent

list filters
  --unread --read --flagged --has-attachments
  --since 30m|6h|3d|2w|YYYY-MM-DD  --until ...
  --sender SUB  --subject SUB  --category C
  --limit N (25)  --scan-max N (2000)  --oldest-first
  --no-preview  --any-class  --count-only

Global
  --pretty      indent the JSON
  --dry-run     mutating commands report the plan and change nothing
  --version     print the version
  --help        this text

Folder paths: "Inbox", "Inbox/Vendors", or store-qualified "projects/inbox".
The English tokens inbox/drafts/sent items/deleted items/junk email/archive work
under any store name and resolve inside that store, so localised mailboxes
(Posteingang, Doruc̎ná pošta) need no special handling.
'@

# ------------------------------------------------------------------- settings

function Get-OlctlHome {
    if (-not [string]::IsNullOrEmpty($env:OLCTL_HOME)) { return $env:OLCTL_HOME }
    $base = $env:USERPROFILE
    if ([string]::IsNullOrEmpty($base)) { $base = $env:HOME }
    return (Join-Path $base '.olctl')
}

function Read-OlctlConfig {
    $cfg = @{
        protect_default_folders = $true
        protected_folders       = @('Deleted Items', 'Junk Email', 'Sent Items', 'Outbox')
        max_items               = 200
        default_body_chars      = 20000
        preview_chars           = 240
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
    $formats = @('yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd')
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

function New-OlctlRestrict {
    param($Opt)
    $clauses = New-Object System.Collections.Generic.List[object]
    if ($Opt.ContainsKey('unread')) { [void]$clauses.Add('[UnRead] = True') }
    if ($Opt.ContainsKey('read')) { [void]$clauses.Add('[UnRead] = False') }
    if ($Opt.ContainsKey('flagged')) { [void]$clauses.Add('[FlagStatus] = 2') }
    if ($Opt.ContainsKey('since')) {
        $w = ConvertTo-OlctlWhen ([string]$Opt['since'])
        [void]$clauses.Add("[ReceivedTime] >= '" + (Get-OlctlRestrictLiteral $w) + "'")
    }
    if ($Opt.ContainsKey('until')) {
        $w = ConvertTo-OlctlWhen ([string]$Opt['until'])
        [void]$clauses.Add("[ReceivedTime] <= '" + (Get-OlctlRestrictLiteral $w) + "'")
    }
    if ($clauses.Count -eq 0) { return $null }
    return [string]::Join(' AND ', $clauses.ToArray())
}

# -------------------------------------------------------------- option parsing

$script:OLCTL_FLAGS = @(
    'unread', 'read', 'flagged', 'has_attachments', 'oldest_first', 'no_preview',
    'any_class', 'count_only', 'headers', 'overwrite', 'all', 'html', 'clear',
    'allow_protected', 'pretty', 'dry_run', 'help', 'h', 'version'
)
$script:OLCTL_VALUES = @(
    'folder', 'shared', 'since', 'until', 'sender', 'subject', 'category', 'limit',
    'scan_max', 'depth', 'body', 'body_chars', 'out', 'pattern', 'due_code', 'text',
    'to', 'cc', 'attach', 'add', 'remove'
)
$script:OLCTL_MULTI = @('to', 'cc', 'attach', 'add', 'remove')

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

function Get-OlctlOptString { param($Opt, [string]$Name, $Default = $null)
    if ($Opt.ContainsKey($Name)) { return [string]$Opt[$Name] } else { return $Default } }

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

# ---------------------------------------------------------------------- main

$script:Envelope = $null
$script:Pretty = $false

try {
    $argv = @()
    if ($null -ne $CliArgs) { $argv = @($CliArgs) }

    $parsed = Get-OlctlParsed -Argv $argv
    $opt = $parsed['opt']
    $pos = $parsed['pos']
    $script:Pretty = $opt.ContainsKey('pretty')
    $dryRun = $opt.ContainsKey('dry_run')

    if ($opt.ContainsKey('version')) {
        $script:Envelope = @{ ok = $true; data = @{ olctl_version = $script:OLCTL_VERSION } }
    } elseif ($opt.ContainsKey('help') -or $opt.ContainsKey('h') -or $pos.Count -eq 0) {
        # Help is for humans, so it is plain text rather than JSON.
        Write-Output $script:USAGE
        exit 0
    } else {
        $cmd = ([string]$pos[0]).ToLowerInvariant()
        $cfg = Read-OlctlConfig
        $a = @{}
        $op = $null
        $auditAs = $null

        switch ($cmd) {
            'doctor' { $op = 'doctor' }

            'folders' {
                $op = 'folders'
                $a['folder'] = Get-OlctlOptString $opt 'folder' $null
                $a['shared'] = Get-OlctlOptString $opt 'shared' $null
                $a['depth'] = Get-OlctlOptInt $opt 'depth' 2
            }

            'list' {
                $op = 'list'
                if (-not $opt.ContainsKey('folder')) {
                    Invoke-OlFail -Code 'bad_argument' -Message 'list needs --folder'
                }
                $maxItems = 200
                try { $maxItems = [int]$cfg['max_items'] } catch { $maxItems = 200 }
                $limit = Get-OlctlOptInt $opt 'limit' 25
                if ($limit -gt $maxItems) { $limit = $maxItems }
                $previewChars = 240
                try { $previewChars = [int]$cfg['preview_chars'] } catch { $previewChars = 240 }
                if ($opt.ContainsKey('no_preview')) { $previewChars = 0 }

                $a['folder'] = Get-OlctlOptString $opt 'folder' $null
                $a['shared'] = Get-OlctlOptString $opt 'shared' $null
                # Built before COM is touched, so a bad --since fails fast.
                $a['restrict'] = New-OlctlRestrict $opt
                $a['limit'] = $limit
                $a['scan_max'] = Get-OlctlOptInt $opt 'scan_max' 2000
                $a['oldest_first'] = $opt.ContainsKey('oldest_first')
                $a['preview_chars'] = $previewChars
                $a['any_class'] = $opt.ContainsKey('any_class')
                $a['count_only'] = $opt.ContainsKey('count_only')
                $a['filters'] = @{
                    sender           = Get-OlctlOptString $opt 'sender' $null
                    subject          = Get-OlctlOptString $opt 'subject' $null
                    category         = Get-OlctlOptString $opt 'category' $null
                    has_attachments  = $opt.ContainsKey('has_attachments')
                }
            }

            'get' {
                $op = 'get'
                $bodyChars = 20000
                try { $bodyChars = [int]$cfg['default_body_chars'] } catch { $bodyChars = 20000 }
                $a['ref'] = Get-OlctlRequiredRef $pos 'get'
                $a['body'] = Get-OlctlOptString $opt 'body' 'text'
                if (@('text', 'html', 'none') -notcontains $a['body']) {
                    Invoke-OlFail -Code 'bad_argument' -Message '--body must be text, html or none'
                }
                $a['body_chars'] = Get-OlctlOptInt $opt 'body_chars' $bodyChars
                $a['headers'] = $opt.ContainsKey('headers')
            }

            'save-attachments' {
                $op = 'save_attachments'
                if (-not $opt.ContainsKey('out')) {
                    Invoke-OlFail -Code 'bad_argument' -Message 'save-attachments needs --out'
                }
                $a['ref'] = Get-OlctlRequiredRef $pos 'save-attachments'
                $a['out'] = Get-OlctlOptString $opt 'out' $null
                $a['pattern'] = Get-OlctlOptString $opt 'pattern' $null
                $a['overwrite'] = $opt.ContainsKey('overwrite')
                $a['dry_run'] = $dryRun
                $auditAs = 'save-attachments'
            }

            'move' {
                $op = 'move'
                if (-not $opt.ContainsKey('to')) {
                    Invoke-OlFail -Code 'bad_argument' -Message 'move needs --to'
                }
                $toList = Get-OlctlOptList $opt 'to'
                $a['ref'] = Get-OlctlRequiredRef $pos 'move'
                $a['to'] = [string]$toList[0]
                $a['shared'] = Get-OlctlOptString $opt 'shared' $null
                $a['allow_protected'] = $opt.ContainsKey('allow_protected')
                $a['protect_defaults'] = [bool]$cfg['protect_default_folders']
                $a['protected_names'] = @($cfg['protected_folders'])
                $a['dry_run'] = $dryRun
                $auditAs = 'move'
            }

            'mark' {
                $op = 'mark'
                $a['ref'] = Get-OlctlRequiredRef $pos 'mark'
                if ($pos.Count -lt 3) {
                    Invoke-OlFail -Code 'bad_argument' -Message 'mark needs read or unread'
                }
                $state = ([string]$pos[2]).ToLowerInvariant()
                if (@('read', 'unread') -notcontains $state) {
                    Invoke-OlFail -Code 'bad_argument' -Message "mark takes 'read' or 'unread'"
                }
                $a['unread'] = ($state -eq 'unread')
                $a['dry_run'] = $dryRun
                $auditAs = 'mark'
            }

            'flag' {
                $op = 'flag'
                $a['ref'] = Get-OlctlRequiredRef $pos 'flag'
                $a['clear'] = $opt.ContainsKey('clear')
                $a['text'] = Get-OlctlOptString $opt 'text' 'Follow up'
                $a['due_code'] = Get-OlctlOptInt $opt 'due_code' 3
                $a['dry_run'] = $dryRun
                $auditAs = 'flag'
            }

            'categorize' {
                $op = 'categorize'
                $a['ref'] = Get-OlctlRequiredRef $pos 'categorize'
                $a['add'] = Get-OlctlOptList $opt 'add'
                $a['remove'] = Get-OlctlOptList $opt 'remove'
                if ($opt.ContainsKey('set')) {
                    $setList = $opt['set']
                    $a['set'] = @($setList.ToArray())
                }
                $a['dry_run'] = $dryRun
                $auditAs = 'categorize'
            }

            'draft-reply' {
                $op = 'draft_reply'
                if (-not $opt.ContainsKey('text')) {
                    Invoke-OlFail -Code 'bad_argument' -Message 'draft-reply needs --text'
                }
                $a['ref'] = Get-OlctlRequiredRef $pos 'draft-reply'
                $a['text'] = Resolve-OlctlText (Get-OlctlOptString $opt 'text' '')
                $a['reply_all'] = $opt.ContainsKey('all')
                $a['html'] = $opt.ContainsKey('html')
                $a['subject'] = Get-OlctlOptString $opt 'subject' $null
                $a['dry_run'] = $dryRun
                $auditAs = 'draft-reply'
            }

            'draft-new' {
                $op = 'draft_new'
                if (-not $opt.ContainsKey('to')) {
                    Invoke-OlFail -Code 'bad_argument' -Message 'draft-new needs at least one --to'
                }
                if (-not $opt.ContainsKey('text')) {
                    Invoke-OlFail -Code 'bad_argument' -Message 'draft-new needs --text'
                }
                $a['to'] = Get-OlctlOptList $opt 'to'
                $a['cc'] = Get-OlctlOptList $opt 'cc'
                $a['subject'] = Get-OlctlOptString $opt 'subject' $null
                $a['text'] = Resolve-OlctlText (Get-OlctlOptString $opt 'text' '')
                $a['html'] = $opt.ContainsKey('html')
                $a['attach'] = Get-OlctlOptList $opt 'attach'
                $a['dry_run'] = $dryRun
                $auditAs = 'draft-new'
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

        $captured = @(Invoke-OlOp -Op $op -OpArgs $opArgs)
        $data = $null
        for ($k = $captured.Count - 1; $k -ge 0; $k--) {
            if ($captured[$k] -is [hashtable]) { $data = $captured[$k]; break }
        }
        if ($null -eq $data) { $data = @{} }

        if (($null -ne $auditAs) -and (-not $dryRun)) {
            Add-OlctlAudit -Action $auditAs -Detail $data
        }

        $envelope = @{}
        $envelope['ok'] = $true
        $envelope['data'] = $data
        $script:Envelope = $envelope
    }
} catch {
    $payload = $null
    $target = $null
    try { $target = $_.TargetObject } catch { $target = $null }
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
        try { $msg = [string]$_.Exception.Message } catch { }
        if ([string]::IsNullOrEmpty($msg)) { $msg = 'unexpected failure' }
        $detail = $null
        try {
            $detail = ($_.Exception.GetType().FullName + ' at line ' +
                [string]$_.InvocationInfo.ScriptLineNumber)
        } catch { $detail = $null }
        $envelope['code'] = 'error'
        $envelope['error'] = $msg
        $envelope['detail'] = $detail
    }
    $script:Envelope = $envelope
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
    # Write-Output, not [Console]::Out.Write. The latter writes straight to the
    # console handle, which bypasses the PowerShell pipeline entirely -- so when
    # the script is dot-invoked in-process (.\olctl.ps1 ... | ConvertFrom-Json)
    # the caller captures nothing and the JSON just lands on screen. Going
    # through the pipeline works both in-process and as a child process.
    try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
    Write-Output $outJson
    exit $exit
} catch {
    try {
        Write-Output '{"ok":false,"code":"error","error":"the response could not be serialised to JSON","detail":null}'
        exit 1
    } catch { exit 3 }
}
