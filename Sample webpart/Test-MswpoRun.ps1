#Requires -Version 7.2
<#
.SYNOPSIS
    Test-site helper for docs/TEST_RUN.md: creates the test operator, switches your MSWPO role, records what the app
    saved (snapshots + change log), runs data checks and cleans up. Refuses to run against production.

.DESCRIPTION
    Actions
      Setup        Section A of TEST_RUN.md: test operator in Operator LookUp, Employee Data, Safety (last + this year),
                   Lab DOC and a Document Set. Remembers every created item and the highest item ID of every MSWPO
                   list (so Cleanup only touches what this run created). Takes a first snapshot.
      Role         Makes you a member of exactly one Access Management group (-Role Management | Supervisors | Trainer |
                   MSWPOAdmin) or of none (-Role Operator). Your original groups are remembered for RestoreRole.
      RestoreRole  Puts your Access Management membership back as it was before the first Role call.
      Snapshot     Reads every item of the test operator in all MSWPO lists (and new station assignment items),
                   saves docs/test-run-data/snapshots/NN-<Label>.json and appends the field-level changes since the
                   previous snapshot to docs/test-run-data/CHANGES.md. Run it after each step: -Label C1.
      Check        -Step A (setup present and correct) | SSA (station assignment rows, reversals, hour sums) |
                   Forms (which feeder/CDSS forms exist for the operator).
      Status       What this run has created so far and your current groups.
      Cleanup      Lists everything this run created; with -Force moves it to the site recycle bin and restores
                   your groups.

    Lists are found like the web part does on a non-production site (decision D20): production GUID, else URL
    name, else title - read from src/webparts/mswpo/config/lists.ts so the two never drift apart.

.REQUIREMENTS
    PowerShell 7.2+, PnP.PowerShell 2.x+, the Entra app ClientId used for Deploy-Mswpo.ps1, edit rights on the test site.

.EXAMPLE
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl https://arlingtonva.sharepoint.com/sites/<test> -Action Setup
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Role -Role Supervisors
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Snapshot -Label C1
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Check -Step SSA
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Cleanup -Force
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ClientId,
    [Parameter(Mandatory)] [string] $SiteUrl,
    [Parameter(Mandatory)] [ValidateSet('Setup', 'Role', 'RestoreRole', 'Snapshot', 'Check', 'Status', 'Cleanup')] [string] $Action,
    [ValidateSet('Management', 'Supervisors', 'Trainer', 'MSWPOAdmin', 'Operator')] [string] $Role,
    [string] $Label,
    [ValidateSet('A', 'SSA', 'Forms')] [string] $Step,
    # Setup values. Default operator = you (so Operator-role steps and "email the operator" reach you).
    [string] $OperatorName,
    [string] $OperatorEmail,
    [string] $Shift = 'A',
    [string] $OpsLicenseLevel = '1',
    [datetime] $HireDate = (Get-Date).Date.AddDays(-180),
    # Safety percentages: face-to-face, online, total
    [double[]] $SafetyLastYear = @(80, 90, 85),
    [double[]] $SafetyThisYear = @(40, 50, 45),
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
$ProductionPath = '/sites/wpcb/mswpo'
$DataDir = Join-Path $PSScriptRoot '..\docs\test-run-data'
$StateFile = Join-Path $DataDir 'state.json'
$SnapDir = Join-Path $DataDir 'snapshots'
$ChangesFile = Join-Path $DataDir 'CHANGES.md'
$ListsTs = Join-Path $PSScriptRoot '..\src\webparts\mswpo\config\lists.ts'

# ---- production guard -----------------------------------------------------------------------------
$sitePath = ([Uri]$SiteUrl).AbsolutePath.TrimEnd('/')
if ($sitePath -ieq $ProductionPath -or $sitePath -eq '') {
    throw "Refusing to run: $SiteUrl is the production MSWPO site. Use the test site URL."
}

# ---- list definitions from lists.ts ---------------------------------------------------------------
function Read-ListConfig {
    $core = [ordered]@{}
    $station = @()
    $kind = ''; $level = ''; $inStation = $false
    foreach ($line in Get-Content $ListsTs) {
        if ($line -match 'export const STATION_LEVEL_LISTS') { $inStation = $true; continue }
        if ($line -match 'export const TRAINING_LIBRARIES') { break }
        if ($inStation -and $line -match '^\s*(cdss|taskStandards|trainingDocs): \{') { $kind = $Matches[1]; continue }
        if ($inStation -and $line -match '^\s*([123]): \{') { $level = $Matches[1]; continue }
        if ($line -match "^\s*(\w+): mswpo\('([0-9A-Fa-f-]{36})', '([^']*)'(?:, '([^']*)')?\)") {
            $def = [ordered]@{ Key = $Matches[1]; Id = $Matches[2]; Title = $Matches[3]; Url = $Matches[4] }
            if ($inStation) {
                $def.Key = "$kind-L$level-$($Matches[1])"
                $def.MatchField = if ($kind -eq 'trainingDocs') { 'Title' } else { 'Employee_x0020_Name' }
                $station += [pscustomobject]$def
            } else {
                $core[$Matches[1]] = [pscustomobject]$def
            }
        }
    }
    if (-not $core.Contains('operatorLookup') -or $station.Count -lt 20) { throw "Could not read the list definitions from $ListsTs" }
    return @{ Core = $core; Station = $station }
}

# ---- state ----------------------------------------------------------------------------------------
function Get-State {
    if (Test-Path $StateFile) { return Get-Content $StateFile -Raw | ConvertFrom-Json -AsHashtable }
    return $null
}
function Save-State($State) {
    New-Item -ItemType Directory -Force $DataDir | Out-Null
    $State | ConvertTo-Json -Depth 8 | Set-Content $StateFile -Encoding utf8
}

# ---- SharePoint helpers ---------------------------------------------------------------------------
$ListCache = @{}
$FieldCache = @{}

function Resolve-MswpoList($Def) {
    if ($ListCache.ContainsKey($Def.Key)) { return $ListCache[$Def.Key] }
    $tries = @(@($Def.Id, 'GUID'))
    if ($Def.Url) { $tries += , @("Lists/$($Def.Url)", 'URL'); $tries += , @($Def.Url, 'URL') }
    $tries += , @($Def.Title, 'title')
    $found = $null
    foreach ($t in $tries) {
        try { $found = Get-PnPList -Identity $t[0] -ErrorAction Stop } catch { $found = $null }
        if ($found) { Write-Verbose "$($Def.Key): found by $($t[1])"; break }
    }
    if (-not $found) { throw "List not found on $SiteUrl : $($Def.Title) (url name '$($Def.Url)')" }
    $ListCache[$Def.Key] = $found
    return $found
}

function Get-FieldInfo($List) {
    $key = [string]$List.Id
    if ($FieldCache.ContainsKey($key)) { return $FieldCache[$key] }
    $info = @{}
    foreach ($f in Get-PnPField -List $List) {
        $x = ([xml]$f.SchemaXml).Field
        $info[$f.InternalName] = [pscustomobject]@{
            Name       = $f.InternalName
            Title      = $f.Title
            Type       = $f.TypeAsString
            Hidden     = $f.Hidden
            ReadOnly   = $f.ReadOnlyField
            Required   = $f.Required
            Choices    = @($x.SelectNodes('.//CHOICE') | ForEach-Object { $_.InnerText })
            LookupList = [string]$x.List
            ShowField  = if ($x.ShowField) { [string]$x.ShowField } else { 'Title' }
            Percentage = ([string]$x.Percentage -ieq 'TRUE')
        }
    }
    $FieldCache[$key] = $info
    return $info
}

function Format-Value($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [Microsoft.SharePoint.Client.FieldLookupValue]) { return [string]$v.LookupValue }
    if ($v -is [Microsoft.SharePoint.Client.FieldUrlValue]) { return [string]$v.Url }
    if ($v -is [datetime]) { return $v.ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
    if ($v -is [System.Array]) { return (($v | ForEach-Object { Format-Value $_ }) -join '; ') }
    return ([string]$v).Trim()
}

function Get-Items($List, [string]$Caml) {
    if ($Caml) { return @(Get-PnPListItem -List $List -Query $Caml) }
    return @(Get-PnPListItem -List $List -PageSize 2000)
}

function Get-MaxId($List) {
    $caml = "<View><Query><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><ViewFields><FieldRef Name='ID'/></ViewFields><RowLimit>1</RowLimit></View>"
    $top = @(Get-PnPListItem -List $List -Query $caml)
    if ($top.Count) { return [int]$top[0].Id }
    return 0
}

function Get-ItemsAbove($List, [int]$MinId) {
    $caml = "<View><Query><Where><Gt><FieldRef Name='ID'/><Value Type='Counter'>$MinId</Value></Gt></Where></Query><RowLimit Paged='TRUE'>500</RowLimit></View>"
    return @(Get-PnPListItem -List $List -Query $caml)
}

function Select-Choice($Field, [string]$Wanted, [string]$What) {
    if (-not $Field.Choices.Count) { return $Wanted }
    $hit = $Field.Choices | Where-Object { $_ -ieq $Wanted } | Select-Object -First 1
    if (-not $hit) { $hit = $Field.Choices | Where-Object { $_ -imatch "(^|\W)$([regex]::Escape($Wanted))(\W|$)" } | Select-Object -First 1 }
    if (-not $hit) { throw "$What '$Wanted' is not a choice of '$($Field.Title)'. Choices: $($Field.Choices -join ', ')" }
    return $hit
}

function Select-ActiveStatus($Field) {
    if (-not $Field) { return 'Active' }
    $active = $Field.Choices | Where-Object { $_ -ieq 'Active' } | Select-Object -First 1
    if (-not $active) { $active = $Field.Choices | Where-Object { $_ -notmatch 'Inactive' } | Select-Object -First 1 }
    if (-not $active) { return 'Active' }
    return $active
}

function Resolve-LookupId($Field, [string]$Text) {
    $target = Get-PnPList -Identity ($Field.LookupList.Trim('{', '}'))
    $match = Get-Items $target | Where-Object { (Format-Value $_.FieldValues[$Field.ShowField]) -eq $Text } | Select-Object -Last 1
    if (-not $match) { throw "No item '$Text' in '$($target.Title)' for lookup '$($Field.Title)'" }
    return [int]$match.Id
}

# Builds the -Values hashtable, converting each wanted value to the field's type; skips fields the list lacks.
function New-Values($List, [System.Collections.IDictionary]$Wanted) {
    $fields = Get-FieldInfo $List
    $values = @{}
    foreach ($name in $Wanted.Keys) {
        $f = $fields[$name]
        $v = $Wanted[$name]
        if (-not $f) { Write-Warning "$($List.Title): field $name not found - skipped"; continue }
        switch -Regex ($f.Type) {
            '^Choice$|^MultiChoice$' { $values[$name] = Select-Choice $f ([string]$v) $name }
            '^Lookup' { $values[$name] = Resolve-LookupId $f ([string]$v) }
            '^Number$|^Currency$' { $values[$name] = if ($f.Percentage) { [double]$v / 100 } else { [double]$v } }
            '^DateTime$' { $values[$name] = [datetime]$v }
            '^Boolean$' { $values[$name] = [bool]$v }
            default { $values[$name] = [string]$v }
        }
    }
    if ($fields['Title'] -and $fields['Title'].Required -and -not $values.ContainsKey('Title')) { $values['Title'] = $OperatorName }
    return $values
}

function Add-TrackedItem($State, [string]$ListKey, $List, $Values) {
    $item = Add-PnPListItem -List $List -Values $Values
    $State.created += , @{ list = $ListKey; id = [int]$item.Id }
    Save-State $State
    Write-Host "  + $($List.Title) #$($item.Id)"
    return $item
}

function Find-OperatorItems($List, [string]$Field, [string]$Name) {
    return @(Get-Items $List | Where-Object { (Format-Value $_.FieldValues[$Field]) -eq $Name })
}

# The document set folder of an operator, by path (the library can be above the 5,000-item threshold).
function Get-DocSetName([string]$Name) { if ($Name.EndsWith('.')) { return $Name.Substring(0, $Name.Length - 1) } return $Name }
function Get-DocSetFolder([string]$Name) {
    $lib = Get-Core 'trainingDocumentSets'
    $root = (Get-PnPProperty -ClientObject $lib -Property RootFolder).ServerRelativeUrl
    $url = "$root/$(Get-DocSetName $Name)"
    try { $folder = Get-PnPFolder -Url $url -Includes ListItemAllFields -ErrorAction Stop } catch { return $null }
    return [pscustomobject]@{ Url = $url; Item = $folder.ListItemAllFields }
}

# ---- connect --------------------------------------------------------------------------------------
$Lists = Read-ListConfig
New-Item -ItemType Directory -Force $DataDir | Out-Null
Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
$Me = Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser
Write-Host "Signed in as $($Me.Title) <$($Me.Email)>"
$State = Get-State

function Get-Core([string]$Key) { return Resolve-MswpoList $Lists.Core[$Key] }

function Get-RunOperator {
    if ($State -and $State.operator) { return $State.operator }
    if ($OperatorName) { return $OperatorName }
    return $Me.Title
}

# ---- roles ----------------------------------------------------------------------------------------
function Get-MemberIds($Item) {
    return @($Item.FieldValues['Group_x0020_Members'] | Where-Object { $_ } | ForEach-Object { [int]$_.LookupId })
}

function Set-Membership($Item, [int[]]$Ids) {
    $values = @($Ids | Select-Object -Unique | ForEach-Object {
            $u = New-Object Microsoft.SharePoint.Client.FieldUserValue
            $u.LookupId = $_
            $u
        })
    $Item['Group_x0020_Members'] = if ($values.Count) { [Microsoft.SharePoint.Client.FieldUserValue[]]$values } else { $null }
    $Item.Update()
    Invoke-PnPQuery
}

function Get-ExpectedRole($Items) {
    # Legacy rule (roleLogic.ts): ID order, last match wins, Supervisors is never replaced.
    $resolved = ''
    foreach ($it in ($Items | Sort-Object { [int]$_.Id })) {
        if ((Get-MemberIds $it) -contains $Me.Id -and $resolved -ne 'Supervisors') { $resolved = [string]$it.FieldValues['Title'] }
    }
    $isOperator = (Find-OperatorItems (Get-Core 'operatorLookup') 'Operator_x0020_Name' $Me.Title).Count -gt 0
    if ($resolved -eq 'Trainer') { if ($isOperator) { return 'Trainer-Operator' } else { return 'Trainer' } }
    if (-not $resolved) { if ($isOperator) { return 'Operator' } else { return 'Visitor' } }
    return $resolved
}

function Invoke-Role([string]$Target) {
    $list = Get-Core 'accessManagement'
    $items = Get-Items $list | Sort-Object { [int]$_.Id }
    $st = $State
    if (-not $st) { $st = @{ created = @() } }
    if (-not $st.originalRoles) {
        $st.originalRoles = @($items | Where-Object { (Get-MemberIds $_) -contains $Me.Id } | ForEach-Object { [int]$_.Id })
        Save-State $st
        Write-Host "Saved your original groups (item ids: $($st.originalRoles -join ', '))."
    }
    $group = if ($Target -eq 'Operator') { '' } else { $Target }
    $targetItem = $items | Where-Object { [string]$_.FieldValues['Title'] -eq $group } | Select-Object -Last 1
    if ($group -and -not $targetItem) { throw "No Access Management item titled '$group'." }
    foreach ($it in $items) {
        $ids = Get-MemberIds $it
        $isMe = $ids -contains $Me.Id
        $want = $targetItem -and [int]$it.Id -eq [int]$targetItem.Id
        if ($want -and -not $isMe) { Set-Membership $it ($ids + $Me.Id); Write-Host "  added to   $($it.FieldValues['Title']) (#$($it.Id))" }
        if (-not $want -and $isMe) { Set-Membership $it ($ids | Where-Object { $_ -ne $Me.Id }); Write-Host "  removed from $($it.FieldValues['Title']) (#$($it.Id))" }
    }
    $role = Get-ExpectedRole (Get-Items $list)
    Write-Host "Role the app should now show: $role  (reload the page)" -ForegroundColor Green
    return $st
}

function Invoke-RestoreRole {
    if (-not $State -or $null -eq $State.originalRoles) { Write-Host 'Nothing to restore (Role was never used in this run).'; return }
    $list = Get-Core 'accessManagement'
    $items = Get-Items $list
    foreach ($it in $items) {
        $ids = Get-MemberIds $it
        $want = $State.originalRoles -contains [int]$it.Id
        $isMe = $ids -contains $Me.Id
        if ($want -and -not $isMe) { Set-Membership $it ($ids + $Me.Id); Write-Host "  added to   $($it.FieldValues['Title'])" }
        if (-not $want -and $isMe) { Set-Membership $it ($ids | Where-Object { $_ -ne $Me.Id }); Write-Host "  removed from $($it.FieldValues['Title'])" }
    }
    $State.Remove('originalRoles')
    Save-State $State
    Write-Host "Groups restored. Role the app should show: $(Get-ExpectedRole (Get-Items $list))" -ForegroundColor Green
}

# ---- setup (section A) ----------------------------------------------------------------------------
function Invoke-Setup {
    if ($State -and $State.created -and $State.created.Count -and -not $Force) {
        throw "A run is already set up for '$($State.operator)' ($($State.created.Count) items). Run -Action Cleanup -Force first (or Setup -Force to add another)."
    }
    $name = if ($OperatorName) { $OperatorName } else { $Me.Title }
    $email = if ($OperatorEmail) { $OperatorEmail } else { $Me.Email }
    $script:OperatorName = $name
    $lookup = Get-Core 'operatorLookup'
    if ((Find-OperatorItems $lookup 'Operator_x0020_Name' $name).Count) { throw "'$name' is already in the Operator LookUp List on this site - pick another -OperatorName or clean up first." }

    $st = if ($State) { $State } else { @{} }
    $st.operator = $name
    $st.email = $email
    $st.startedAt = (Get-Date).ToString('o')
    if (-not $st.created) { $st.created = @() }
    Write-Host "Recording the highest item id of every MSWPO list ..."
    $st.baseline = @{}
    foreach ($key in $Lists.Core.Keys) { $st.baseline[$key] = Get-MaxId (Get-Core $key) }
    foreach ($def in $Lists.Station) { $st.baseline[$def.Key] = Get-MaxId (Resolve-MswpoList $def) }
    Save-State $st

    Write-Host "Creating test operator '$name' <$email>, shift $Shift" -ForegroundColor Cyan
    # A1 Operator LookUp
    $a1 = [ordered]@{
        Operator_x0020_Name = $name; Operator_x0020_Email = $email; Operator_x0020_Shift = $Shift; Job_x0020_Title = 'Operator'
    }
    $lf = Get-FieldInfo $lookup
    $a1.Status = Select-ActiveStatus $lf['Status']
    'PrimaryStationHours', 'PrimaryStationTrainingHours', 'SecondaryStationHours', 'SecondaryStationTrainingHours', 'AWTStationHours',
    'AWTStationTrainingHours', 'DWBStationHours', 'DWBStationTrainingHours', 'SupervisorDutiesStationHours',
    'SupervisorDutiesStationTrainingHours', 'SupervisorDutiesHours' | Where-Object { $lf[$_] } | ForEach-Object { $a1[$_] = 0 }
    Add-TrackedItem $st 'operatorLookup' $lookup (New-Values $lookup $a1) | Out-Null

    # Shift supervisor = legacy GetSupervisorEmail: active Supervisor on the same shift, last match wins.
    $supShift = Select-Choice $lf['Operator_x0020_Shift'] $Shift 'Shift'
    $sup = Get-Items $lookup | Where-Object {
        (Format-Value $_.FieldValues['Operator_x0020_Shift']) -eq $supShift -and
        (Format-Value $_.FieldValues['Job_x0020_Title']) -eq 'Supervisor' -and
        (Format-Value $_.FieldValues['Status']) -ne 'Inactive'
    } | Select-Object -Last 1

    # A2 Employee Data
    $emp = Get-Core 'operatorCertification'
    $ef = Get-FieldInfo $emp
    $a2 = [ordered]@{
        System_x0020_Defined_x0020_Opera = $name; Hire_x0020_Date = $HireDate; OPS_x0020_License_x0020_Level = $OpsLicenseLevel; Shift = $Shift
        Primary_x0020_Current_x0020_Leve = 'T'; Secondary_x0020_Current_x0020_Le = 'T'; AWT_x0020_Current_x0020_Level = 'T'; DWB_x0020_Current_x0020_Level = 'T'
    }
    $a2.Status = Select-ActiveStatus $ef['Status']
    if ($sup -and $ef['Shift_x0020_Supervisor']) {
        $a2.Shift_x0020_Supervisor = if ($ef['Shift_x0020_Supervisor'].Type -like 'User*') { Format-Value $sup.FieldValues['Operator_x0020_Email'] } else { Format-Value $sup.FieldValues['Operator_x0020_Name'] }
    }
    Add-TrackedItem $st 'operatorCertification' $emp (New-Values $emp $a2) | Out-Null

    # A3 Safety, last year and this year
    $safety = Get-Core 'safetyTraining'
    $sf = Get-FieldInfo $safety
    $year = (Get-Date).Year
    foreach ($row in @(@($year - 1, $SafetyLastYear), @($year, $SafetyThisYear))) {
        $y = if ($sf['Year'] -and $sf['Year'].Type -match 'Number') { [int]$row[0] } else { [string]$row[0] }
        $a3 = [ordered]@{
            Operator_x0020_Look_x0020_up = $name; Year = $y
            _x0025_Face_x0020_to_x0020_Face_ = $row[1][0]; _x0025_Online_x0020_Completed = $row[1][1]; _x0025__x0020_Total_x0020_Safety = $row[1][2]
        }
        Add-TrackedItem $st 'safetyTraining' $safety (New-Values $safety $a3) | Out-Null
    }

    # A4 Lab DOC (dates blank)
    $doc = Get-Core 'docCertificationDates'
    $df = Get-FieldInfo $doc
    $a4 = [ordered]@{ Title = $name; Shift = $Shift }
    if ($df['Status']) { $a4.Status = Select-ActiveStatus $df['Status'] }
    Add-TrackedItem $st 'docCertificationDates' $doc (New-Values $doc $a4) | Out-Null

    # A5 Document set (name without a trailing dot - legacy GetDocumentSetURL), Title = operator name
    $lib = Get-Core 'trainingDocumentSets'
    $dsName = Get-DocSetName $name
    if (Get-DocSetFolder $name) { throw "A document set '$dsName' already exists in '$($lib.Title)'." }
    $ct = Get-PnPContentType -List $lib | Where-Object { $_.Id.StringValue -like '0x0120D520*' } | Select-Object -First 1
    if (-not $ct) { throw "No Document Set content type on '$($lib.Title)'." }
    Add-PnPDocumentSet -List $lib -ContentType $ct.Name -Name $dsName | Out-Null
    $folder = Get-DocSetFolder $name
    if (-not $folder) { throw "Document set '$dsName' was not found after creating it." }
    Set-PnPListItem -List $lib -Identity $folder.Item.Id -Values @{ Title = $name } | Out-Null
    $st.created += , @{ list = 'trainingDocumentSets'; id = [int]$folder.Item.Id }
    Save-State $st
    Write-Host "  + $($lib.Title) document set '$dsName' #$($folder.Item.Id)"
    return $st
}

# ---- snapshot & changes ---------------------------------------------------------------------------
$SkipFields = @('ContentType', 'Attachments', 'Edit', 'LinkTitle', 'LinkTitleNoMenu', 'DocIcon', 'ItemChildCount', 'FolderChildCount',
    '_UIVersionString', 'AppAuthor', 'AppEditor', '_ComplianceFlags', '_ComplianceTag', '_ComplianceTagWrittenTime',
    '_ComplianceTagUserId', '_IsRecord', 'SelectTitle', 'LinkFilename', 'LinkFilenameNoMenu', '_CheckinComment', 'CheckoutUser')
$KeepFields = @('ID', 'Created', 'Modified', 'Author', 'Editor', 'FileLeafRef')

function ConvertTo-Row($Item, $Fields) {
    $row = [ordered]@{}
    foreach ($f in $Fields.Values | Sort-Object Name) {
        if ($SkipFields -contains $f.Name) { continue }
        if ($f.Hidden -and $KeepFields -notcontains $f.Name) { continue }
        if ($f.Type -eq 'Computed' -and $KeepFields -notcontains $f.Name) { continue }
        $v = Format-Value $Item.FieldValues[$f.Name]
        if ($v -ne '' -or $KeepFields -contains $f.Name) { $row[$f.Name] = $v }
    }
    return $row
}

function Get-Snapshot {
    $name = Get-RunOperator
    $snap = [ordered]@{}
    $add = {
        param($Key, $List, $Items)
        $fields = Get-FieldInfo $List
        $rows = [ordered]@{}
        foreach ($it in ($Items | Sort-Object { [int]$_.Id })) { $rows[[string]$it.Id] = ConvertTo-Row $it $fields }
        $snap[$Key] = [ordered]@{ title = $List.Title; items = $rows }
    }
    & $add 'operatorLookup' (Get-Core 'operatorLookup') (Find-OperatorItems (Get-Core 'operatorLookup') 'Operator_x0020_Name' $name)
    & $add 'operatorCertification' (Get-Core 'operatorCertification') (Find-OperatorItems (Get-Core 'operatorCertification') 'System_x0020_Defined_x0020_Opera' $name)
    & $add 'safetyTraining' (Get-Core 'safetyTraining') (Find-OperatorItems (Get-Core 'safetyTraining') 'Operator_x0020_Look_x0020_up' $name)
    & $add 'docCertificationDates' (Get-Core 'docCertificationDates') (Find-OperatorItems (Get-Core 'docCertificationDates') 'Title' $name)
    $lib = Get-Core 'trainingDocumentSets'
    $set = Get-DocSetFolder $name
    $docs = @()
    if ($set) { $docs = @($set.Item) + @(Get-PnPListItem -List $lib -FolderServerRelativeUrl $set.Url -PageSize 500) }
    & $add 'trainingDocumentSets' $lib $docs
    & $add 'operatorStationAssignment' (Get-Core 'operatorStationAssignment') (Find-OperatorItems (Get-Core 'operatorStationAssignment') 'operator' $name)
    if ($State -and $State.baseline) {
        & $add 'supervisorStationAssignment' (Get-Core 'supervisorStationAssignment') (Get-ItemsAbove (Get-Core 'supervisorStationAssignment') $State.baseline['supervisorStationAssignment'])
    }
    $access = Get-Core 'accessManagement'
    & $add 'accessManagement' $access (Get-Items $access | Where-Object { (Get-MemberIds $_) -contains $Me.Id })
    foreach ($def in $Lists.Station) {
        $list = Resolve-MswpoList $def
        $items = Find-OperatorItems $list $def.MatchField $name
        if ($items.Count) { & $add $def.Key $list $items }
    }
    return $snap
}

function Format-Row($Row) {
    return (($Row.Keys | Where-Object { @('Created', 'Modified', 'Author', 'Editor') -notcontains $_ } | ForEach-Object { "$_=$($Row[$_])" }) -join '; ')
}

function Compare-Snapshot($Prev, $Curr) {
    $lines = @()
    $keys = @($Curr.Keys) + @(if ($Prev) { $Prev.Keys | Where-Object { -not $Curr.Contains($_) } })
    foreach ($k in $keys) {
        $c = if ($Curr.Contains($k)) { $Curr[$k].items } else { @{} }
        $p = if ($Prev -and $Prev.Contains($k)) { $Prev[$k].items } else { @{} }
        $title = if ($Curr.Contains($k)) { $Curr[$k].title } else { $Prev[$k].title }
        $out = @()
        foreach ($id in $c.Keys) {
            if (-not $p.Contains($id)) { $out += "- **added** #${id} (by $($c[$id].Author)): $(Format-Row $c[$id])"; continue }
            $diff = @()
            foreach ($f in (@($c[$id].Keys) + @($p[$id].Keys) | Select-Object -Unique)) {
                if (@('Modified', 'Editor') -contains $f) { continue }
                $old = [string]$p[$id][$f]; $new = [string]$c[$id][$f]
                if ($old -ne $new) { $diff += "${f}: '$old' -> '$new'" }
            }
            if ($diff.Count) { $out += "- **changed** #${id} (by $($c[$id].Editor) at $($c[$id].Modified)): $($diff -join '; ')" }
        }
        foreach ($id in $p.Keys) { if (-not $c.Contains($id)) { $out += "- **removed** #${id}" } }
        if ($out.Count) { $lines += "### $title"; $lines += $out }
    }
    return $lines
}

function Invoke-Snapshot {
    if (-not $Label) { throw 'Give the step as -Label (e.g. -Label C1).' }
    New-Item -ItemType Directory -Force $SnapDir | Out-Null
    $files = @(Get-ChildItem $SnapDir -Filter '*.json' | Sort-Object Name)
    $prev = if ($files.Count) { Get-Content $files[-1].FullName -Raw | ConvertFrom-Json -AsHashtable } else { $null }
    $curr = Get-Snapshot
    $file = Join-Path $SnapDir ('{0:D2}-{1}.json' -f ($files.Count + 1), ($Label -replace '[^\w.-]', '_'))
    $curr | ConvertTo-Json -Depth 8 | Set-Content $file -Encoding utf8
    $role = Get-ExpectedRole (Get-Items (Get-Core 'accessManagement'))
    $lines = Compare-Snapshot $prev (Get-Content $file -Raw | ConvertFrom-Json -AsHashtable)
    if (-not $lines.Count) { $lines = @('_No changes since the previous snapshot._') }
    $header = "## $Label - $(Get-Date -Format 'yyyy-MM-dd HH:mm') (role: $role, operator: $(Get-RunOperator))"
    if (-not (Test-Path $ChangesFile)) { "# Test run - data changes per step`n`nWritten by tools/Test-MswpoRun.ps1 -Action Snapshot. Do not edit by hand.`n" | Set-Content $ChangesFile -Encoding utf8 }
    Add-Content $ChangesFile -Encoding utf8 -Value (@('', $header, '') + $lines)
    Write-Host $header -ForegroundColor Cyan
    $lines | ForEach-Object { Write-Host $_ }
    Write-Host "Saved $file; appended to $ChangesFile"
}

# ---- checks ---------------------------------------------------------------------------------------
function Write-Check([bool]$Ok, [string]$Text) {
    if ($Ok) { Write-Host "  PASS  $Text" -ForegroundColor Green } else { Write-Host "  FAIL  $Text" -ForegroundColor Red }
    $script:CheckLines += "- $(if ($Ok) { 'PASS' } else { '**FAIL**' }) $Text"
}

function Invoke-Check {
    if (-not $Step) { throw 'Give -Step A | SSA | Forms.' }
    $name = Get-RunOperator
    $script:CheckLines = @()
    Write-Host "Check $Step for '$name'" -ForegroundColor Cyan
    switch ($Step) {
        'A' {
            $l = Find-OperatorItems (Get-Core 'operatorLookup') 'Operator_x0020_Name' $name
            Write-Check ($l.Count -eq 1) "Operator LookUp: exactly one item ($($l.Count))"
            if ($l.Count) { Write-Check ((Format-Value $l[0].FieldValues['Status']) -ne 'Inactive') "Operator LookUp: status '$(Format-Value $l[0].FieldValues['Status'])' is not Inactive" }
            $e = Find-OperatorItems (Get-Core 'operatorCertification') 'System_x0020_Defined_x0020_Opera' $name
            Write-Check ($e.Count -eq 1) "Employee Data: exactly one item ($($e.Count))"
            if ($e.Count) {
                foreach ($f in 'Primary_x0020_Current_x0020_Leve', 'Secondary_x0020_Current_x0020_Le', 'AWT_x0020_Current_x0020_Level', 'DWB_x0020_Current_x0020_Level') {
                    Write-Check ((Format-Value $e[0].FieldValues[$f]) -eq 'T') "Employee Data: $f = T ($(Format-Value $e[0].FieldValues[$f]))"
                }
                Write-Check ((Format-Value $e[0].FieldValues['Shift_x0020_Supervisor']) -ne '') "Employee Data: shift supervisor set ($(Format-Value $e[0].FieldValues['Shift_x0020_Supervisor']))"
            }
            $s = Find-OperatorItems (Get-Core 'safetyTraining') 'Operator_x0020_Look_x0020_up' $name
            Write-Check ($s.Count -ge 1) "Safety: $($s.Count) year(s): $(($s | ForEach-Object { Format-Value $_.FieldValues['Year'] }) -join ', ')"
            $d = Find-OperatorItems (Get-Core 'docCertificationDates') 'Title' $name
            Write-Check ($d.Count -eq 1) "Lab DOC: exactly one item ($($d.Count))"
            $set = Get-DocSetFolder $name
            Write-Check ([bool]$set) "Document Set folder '$(Get-DocSetName $name)' exists"
            if ($set) { Write-Check ((Format-Value $set.Item.FieldValues['Title']) -eq $name) "Document Set Title = '$name' (the app finds it by Title) - is '$(Format-Value $set.Item.FieldValues['Title'])'" }
        }
        'SSA' {
            $rows = Find-OperatorItems (Get-Core 'operatorStationAssignment') 'operator' $name | Sort-Object { [int]$_.Id }
            if (-not $rows.Count) { Write-Check $false 'Operator Station Assignment: no rows for the operator' }
            $rows | ForEach-Object {
                [pscustomobject]@{
                    ID = $_.Id; Key = Format-Value $_.FieldValues['Title']; Station = Format-Value $_.FieldValues['Station']; Hours = Format-Value $_.FieldValues['hours']
                    Level = Format-Value $_.FieldValues['level']; Training = Format-Value $_.FieldValues['training']; Recert = Format-Value $_.FieldValues['recertification']
                    Trainer = Format-Value $_.FieldValues['trainerorsme']; OpShift = Format-Value $_.FieldValues['operatorshift']; Shift = Format-Value $_.FieldValues['assignmentshift']
                }
            } | Format-Table -AutoSize | Out-String | Write-Host
            foreach ($g in ($rows | Group-Object { Format-Value $_.FieldValues['Title'] })) {
                $sum = ($g.Group | ForEach-Object { [double](Format-Value $_.FieldValues['hours']) } | Measure-Object -Sum).Sum
                $neg = @($g.Group | Where-Object { [double](Format-Value $_.FieldValues['hours']) -lt 0 }).Count
                Write-Check ($sum -ge 0) "$($g.Name): $($g.Count) row(s), $neg reversal(s), net hours $sum"
            }
            if ($State -and $State.baseline) {
                $ssa = Get-ItemsAbove (Get-Core 'supervisorStationAssignment') $State.baseline['supervisorStationAssignment']
                foreach ($i in $ssa) { Write-Host "  SSA item #$($i.Id): $(Format-Value $i.FieldValues['Title'])" }
            }
        }
        'Forms' {
            foreach ($def in $Lists.Station) {
                $items = Find-OperatorItems (Resolve-MswpoList $def) $def.MatchField $name
                if ($items.Count) {
                    Write-Check ($items.Count -eq 1) "$($def.Key): $(($items | ForEach-Object { "#$($_.Id) modified $(Format-Value $_.FieldValues['Modified']) by $(Format-Value $_.FieldValues['Editor'])" }) -join ' | ')"
                }
            }
            if (-not $script:CheckLines.Count) { Write-Host '  No feeder/CDSS forms yet.' }
        }
    }
    Add-Content $ChangesFile -Encoding utf8 -Value (@('', "## Check $Step - $(Get-Date -Format 'yyyy-MM-dd HH:mm')", '') + $script:CheckLines)
}

# ---- status & cleanup -----------------------------------------------------------------------------
function Get-CleanupItems {
    $name = Get-RunOperator
    $todo = @()
    foreach ($c in $State.created) { $todo += [pscustomobject]@{ Key = $c.list; Id = [int]$c.id; Why = 'created by Setup' } }
    foreach ($def in $Lists.Station) {
        $min = [int]$State.baseline[$def.Key]
        foreach ($i in (Find-OperatorItems (Resolve-MswpoList $def) $def.MatchField $name | Where-Object { [int]$_.Id -gt $min })) {
            $todo += [pscustomobject]@{ Key = $def.Key; Id = [int]$i.Id; Why = 'form saved during the run' }
        }
    }
    foreach ($k in 'operatorStationAssignment', 'supervisorStationAssignment') {
        foreach ($i in (Get-ItemsAbove (Get-Core $k) ([int]$State.baseline[$k]) | Where-Object { $_.FieldValues['Author'].LookupId -eq $Me.Id })) {
            $todo += [pscustomobject]@{ Key = $k; Id = [int]$i.Id; Why = 'station assignment saved by you during the run' }
        }
    }
    # Dependants first, Operator LookUp last (other lists look it up).
    $order = @{ operatorLookup = 9; operatorCertification = 8; safetyTraining = 7; docCertificationDates = 6; trainingDocumentSets = 5 }
    return $todo | Sort-Object { if ($order.ContainsKey($_.Key)) { $order[$_.Key] } else { 1 } }, Key, Id -Unique
}

function Get-DefByKey([string]$Key) {
    if ($Lists.Core.Contains($Key)) { return $Lists.Core[$Key] }
    return $Lists.Station | Where-Object { $_.Key -eq $Key } | Select-Object -First 1
}

function Invoke-Cleanup {
    if (-not $State -or -not $State.baseline) { throw 'No run state found (docs/test-run-data/state.json) - nothing to clean up.' }
    $todo = Get-CleanupItems
    Write-Host "Items of this run for '$(Get-RunOperator)':" -ForegroundColor Cyan
    $todo | ForEach-Object { [pscustomobject]@{ List = (Resolve-MswpoList (Get-DefByKey $_.Key)).Title; ID = $_.Id; Why = $_.Why } } | Format-Table -AutoSize | Out-String | Write-Host
    if (-not $Force) { Write-Host 'Dry run. Add -Force to move these items to the site recycle bin and restore your groups.' -ForegroundColor Yellow; return }
    foreach ($t in $todo) {
        $list = Resolve-MswpoList (Get-DefByKey $t.Key)
        try { Remove-PnPListItem -List $list -Identity $t.Id -Recycle -Force | Out-Null; Write-Host "  recycled $($list.Title) #$($t.Id)" }
        catch { Write-Warning "  $($list.Title) #$($t.Id): $($_.Exception.Message)" }
    }
    Invoke-RestoreRole
    $archive = Join-Path $DataDir ("state-{0:yyyyMMdd-HHmm}.json" -f (Get-Date))
    Move-Item $StateFile $archive
    Write-Host "Run closed; state archived to $archive" -ForegroundColor Green
}

function Invoke-Status {
    if (-not $State) { Write-Host 'No run set up yet (use -Action Setup).'; return }
    Write-Host "Operator: $($State.operator) <$($State.email)>  started $($State.startedAt)"
    Write-Host "Created by Setup: $(($State.created | ForEach-Object { "$($_.list)#$($_.id)" }) -join ', ')"
    if ($null -ne $State.originalRoles) { Write-Host "Original groups (Access Management item ids): $($State.originalRoles -join ', ')" }
    Write-Host "Role the app should show now: $(Get-ExpectedRole (Get-Items (Get-Core 'accessManagement')))"
    $n = if (Test-Path $SnapDir) { @(Get-ChildItem $SnapDir -Filter '*.json').Count } else { 0 }
    Write-Host "Snapshots: $n"
}

# ---- run ------------------------------------------------------------------------------------------
switch ($Action) {
    'Setup' {
        $State = Invoke-Setup
        $Label = 'A-setup'
        Invoke-Snapshot
        $Step = 'A'
        Invoke-Check
    }
    'Role' {
        if (-not $Role) { throw 'Give -Role Management | Supervisors | Trainer | MSWPOAdmin | Operator.' }
        $State = Invoke-Role $Role
    }
    'RestoreRole' { Invoke-RestoreRole }
    'Snapshot' { Invoke-Snapshot }
    'Check' { Invoke-Check }
    'Status' { Invoke-Status }
    'Cleanup' { Invoke-Cleanup }
}
