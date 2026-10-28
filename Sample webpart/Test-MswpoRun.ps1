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
                   Forms (which feeder/CDSS forms exist for the operator) | Lookups (lookup columns whose target
                   list does not exist on this site - lists copied from production keep production list ids).
      Status       What this run has created so far and your current groups.
      Cleanup      Lists everything this run created; with -Force moves it to the site recycle bin and restores
                   your groups.
      FixLookups   Lists the broken lookup columns (Check Lookups); with -Force re-points them to the test copies.
                   Changes the test site's list schema only.

    Lists are found like the web part does on a non-production site (decision D20): production GUID, else URL
    name, else title. The definitions are embedded below (copied from src/webparts/mswpo/config/lists.ts and kept
    in sync by tools/Run-LogicTests.ps1), so this single file can be copied to and run on any machine.

    Output (state, snapshots, CHANGES.md) goes to -DataDir; default: docs/test-run-data when run from the project,
    otherwise a "test-run-data" folder next to this script. Bring CHANGES.md back (or paste the console output).

.REQUIREMENTS
    PowerShell 7.2+, PnP.PowerShell 2.x+, the Entra app ClientId used for Deploy-Mswpo.ps1, edit rights on the test site.
    Sign-in: browser pop-up (-Interactive, like Deploy-Mswpo.ps1); on a server without a usable browser add
    -DeviceLogin and complete the code on any device (add -Tenant <tenant>.onmicrosoft.com if PnP asks for it).

.EXAMPLE
    pwsh ./Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl https://arlingtonva.sharepoint.com/sites/<test> -Action Setup
    pwsh ./Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Setup -DeviceLogin
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Role -Role Supervisors
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Snapshot -Label C1
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Check -Step SSA
    pwsh ./tools/Test-MswpoRun.ps1 -ClientId <guid> -SiteUrl <test site> -Action Cleanup -Force
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ClientId,
    [Parameter(Mandatory)] [string] $SiteUrl,
    [Parameter(Mandatory)] [ValidateSet('Setup', 'Role', 'RestoreRole', 'Snapshot', 'Check', 'Status', 'Cleanup', 'FixLookups')] [string] $Action,
    [ValidateSet('Management', 'Supervisors', 'Trainer', 'MSWPOAdmin', 'Operator')] [string] $Role,
    [string] $Label,
    [ValidateSet('A', 'SSA', 'Forms', 'Lookups')] [string] $Step,
    # Setup values. Default operator = you (so Operator-role steps and "email the operator" reach you).
    [string] $OperatorName,
    [string] $OperatorEmail,
    [string] $Shift = 'A',
    [string] $OpsLicenseLevel = '1',
    [datetime] $HireDate = (Get-Date).Date.AddDays(-180),
    # Safety percentages: face-to-face, online, total
    [double[]] $SafetyLastYear = @(80, 90, 85),
    [double[]] $SafetyThisYear = @(40, 50, 45),
    [switch] $Force,
    [string] $DataDir,
    # Role / RestoreRole: the account to change (email), e.g. the one signed in to the browser. Default: you.
    [string] $User,
    [switch] $DeviceLogin,
    [string] $Tenant
)

$ErrorActionPreference = 'Stop'
$ProductionPath = '/sites/wpcb/mswpo'
if (-not $DataDir) {
    $projectDocs = Join-Path $PSScriptRoot '..\docs'
    $DataDir = if (Test-Path $projectDocs) { Join-Path $projectDocs 'test-run-data' } else { Join-Path $PSScriptRoot 'test-run-data' }
}
$StateFile = Join-Path $DataDir 'state.json'
$SnapDir = Join-Path $DataDir 'snapshots'
$ChangesFile = Join-Path $DataDir 'CHANGES.md'

# BEGIN LISTS - generated from src/webparts/mswpo/config/lists.ts by tools/Run-LogicTests.ps1; do not edit by hand.
$EmbeddedLists = @'
export const CORE_LISTS = {
accessManagement: mswpo('B441CA50-2EFB-425F-BF1B-E66EE18806B0', 'MSWPO Access Management Configuration', 'MSWPO Access Management Configuration'),
operatorLookup: mswpo('F032B017-5260-492F-B72A-1A7E7FC31A23', 'Operator LookUp List', 'Operator LookUp List'),
operatorCertification: mswpo('68C451AE-426B-4CFA-B2DC-B101D8D115DF', 'Operator Certification and Employee Data', 'Employee Data'),
safetyTraining: mswpo('9E08D4B1-10E4-4E14-8E02-31C3379732C0', 'MSWPO Safety Training Completion List', 'MSWPO Safety Training Completion List'),
docCertificationDates: mswpo('D088C1C4-D8DD-4281-9E5C-62D2645CB84A', 'DOC Certification Dates for Lab', 'DOC Certification Dates for Lab'),
trainingDocumentSets: mswpo('C8746D73-8405-44C5-BCE2-0DD0E8CDC332', 'All MSWPO Training Documents', 'All MSWPO Training Documents'),
supervisorStationAssignment: mswpo('C18CC7CF-4BB6-4D55-BCCD-9D252976D75B', 'Supervisor Station Assignment', 'Supervisor Station Assignment'),
operatorStationAssignment: mswpo('C2932C64-DF7B-4221-85D3-5E3A11FDC2D5', 'Operator Station Assignment')
export const STATION_LEVEL_LISTS: Record<FormKind, Record<Level, StationLists>> = {
cdss: {
3: {
primary: mswpo('3b8c1d2c-1839-4491-87b0-30c4926361d5', 'CDSS Primary Station Training Level 3', 'CDSS Primary Station Training Level 3'),
secondary: mswpo('a606bc0f-0f4c-4654-bd1c-e4fca32154ab', 'CDSS- Secondary Station - Training Level 3', 'CDSS Secondary Station  Training Level 3'),
awt: mswpo('a1d23b11-4eb9-40d8-9e3e-565e1b7ee422', 'CDSS- AWT Station - Training Level 3', 'CDSS AWT Station  Training Level 3'),
dwb: mswpo('64d28fce-ff0e-46b4-8c64-76fa27162a04', 'CDSS- DWB Station - Training Level 3', 'CDSS DWB Station  Training Level 3')
2: {
primary: mswpo('99fe0336-193e-42b6-83aa-ff7cc8f44183', 'CDSS Primary Station Training Level 2', 'CDSS Primary Station Training Level 2'),
secondary: mswpo('D8F44610-ED6A-4063-BFCD-2C47339A179A', 'CDSS Secondary Station Training Level 2', 'CDSS Secondary Station Training Level 2'),
awt: mswpo('AF8406F9-2ADB-4617-8FBB-EFA63275F563', 'CDSS AWT Station Training Level 2', 'CDSS AWT Station Training Level 2'),
dwb: mswpo('8F4B7AA0-9E38-4207-AA8B-B6B379289B1D', 'CDSS DWB Station Training Level 2', 'CDSS DWB Station Training Level 2')
1: {
all: mswpo('FCBAA45D-439A-432B-B356-E8A3290BE6FC', 'CDSS- All Stations - Training Level 1', 'CDSS All Stations  Training Level 1')
taskStandards: {
3: {
primary: mswpo('DBEC4BE4-3A1E-403F-B5C5-79340F48445C', 'Primary Station Level 3 Task Standards - RRCLs', 'Primary Station Level 3 Task Standards  RRCLs'),
secondary: mswpo('f0b13014-aba0-4f79-b21d-b9d48d8d5a24', 'Secondary Station Level 3 Task Standards - RRCLs', 'Secondary Station Level 3 Task Standards  RRCLs'),
awt: mswpo('d6e8defa-b6d6-4654-a7af-1f244599c1bd', 'AWT Station Level 3 Task Standards - RRCLs', 'AWT Station Level 3 Task Standards  RRCLs'),
dwb: mswpo('ea0b10c3-cb86-4e19-ab76-f16b55bac9e0', 'DWB Station Level 3 Task Standards - RRCLs', 'DWB Station Level 3 Task Standards  RRCLs')
2: {
primary: mswpo('c5588cd5-0cc4-4b2d-aac7-2281aa41c353', 'Primary Station Level 2 Task Standards - MiniMods', 'Primary Station Level 2 Task Standards  MiniMods'),
secondary: mswpo('8918815a-b12c-4741-b024-957895254cce', 'Secondary Station Level 2 Task Standards - MiniMods', 'Secondary Station Level 2 Task Standards  MiniMods'),
awt: mswpo('2d185a32-a228-4fb7-8e31-e43b20666ac4', 'AWT Station Level 2 Task Standards - MiniMods', 'AWT Station Level 2 Task Standards  MiniMods'),
dwb: mswpo('13e69c0a-f5c5-4e61-8f59-ce5c1221725d', 'DWB Station Level 2 Task Standards - MiniMods', 'DWB Station Level 2 Task Standards  MiniMods')
1: {
all: mswpo('FB0FD350-3EF9-45E4-A293-519434B7AA65', 'All Stations Level 1 Task Standards - Test', 'All Stations Level1 Task Standards  Test')
trainingDocs: {
3: {
primary: mswpo('8a44d44b-8da7-40e7-a057-d94e2954c712', 'Primary Station - Level 3 Feeder Training Docs', 'Primary Station  Level 3 Feeder Training Docs'),
secondary: mswpo('9a113dca-97ec-4c37-86ca-01f95a0d1498', 'Secondary Station - Level 3 Feeder Training Docs', 'Secondary Station  Level 3 Feeder Training Docs'),
awt: mswpo('d9fc725b-9394-46ef-b94e-90bedd717865', 'AWT Station - Level 3 Feeder Training Docs', 'AWT Station  Level 3 Feeder Training Docs'),
dwb: mswpo('5cc83d4e-892d-4f2b-8a7f-4986d74fd68b', 'DWB Station - Level 3 Feeder Training Docs', 'DWB Station  Level 3 Feeder Training Docs')
2: {
primary: mswpo('149e1017-cf92-45f2-b195-fe12a570f8f4', 'Primary Station - Level 2 Feeder Training Docs', 'Primary Station  Level 2 Feeder Training Docs'),
secondary: mswpo('e12a45d2-8faa-43cd-847b-f04b8657a7a9', 'Secondary Station - Level 2 Feeder Training Docs', 'Secondary Station  Level 2 Feeder Training Docs'),
awt: mswpo('b65f8db5-051d-4cfc-b4fb-f33bae7084a0', 'AWT Station - Level 2 Feeder Training Docs', 'AWT Station  Level 2 Feeder Training Docs'),
dwb: mswpo('3674435a-0c3a-4a45-8c8d-81eb98335d13', 'DWB Station - Level 2 Feeder Training Docs', 'DWB Station  Level 2 Feeder Training Docs')
1: {
all: mswpo('9428C900-C93D-4428-A3F9-33F15864672B', 'All Stations - Level 1 Feeder Training Docs', 'All Stations  Level 1 Feeder Training Docs')
export const TRAINING_LIBRARIES: ITrainingLibrary[] = [
'@
# END LISTS

# ---- production guard -----------------------------------------------------------------------------
$sitePath = ([Uri]$SiteUrl).AbsolutePath.TrimEnd('/')
if ($sitePath -ieq $ProductionPath -or $sitePath -eq '') {
    throw "Refusing to run: $SiteUrl is the production MSWPO site. Use the test site URL."
}

# ---- list definitions (embedded copy of lists.ts) -------------------------------------------------
function Read-ListConfig([string[]]$Lines = ($EmbeddedLists -split "`r?`n")) {
    $core = [ordered]@{}
    $station = @()
    $kind = ''; $level = ''; $inStation = $false
    foreach ($line in $Lines) {
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
    if (-not $core.Contains('operatorLookup') -or $station.Count -lt 20) { throw 'Could not read the embedded list definitions (run tools/Run-LogicTests.ps1 in the project to regenerate them).' }
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

# Target list of a lookup column. A list copied from production keeps the PRODUCTION list id in its lookup
# columns ("Broken"): then the test copy of that production list is returned via the embedded definitions.
$LookupTargetCache = @{}
function Find-LookupTarget([string]$ListId) {
    $id = $ListId.Trim('{', '}')
    if ($id -notmatch '^[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}$') { return $null }   # Self, UserInfo, AppPrincipals
    if ($LookupTargetCache.ContainsKey($id)) { return $LookupTargetCache[$id] }
    try { $l = Get-PnPList -Identity $id -ErrorAction Stop } catch { $l = $null }
    if ($l) {
        $t = [pscustomobject]@{ List = $l; Broken = $false; Def = $null }
    } else {
        $def = @($Lists.Core.Values) + @($Lists.Station) | Where-Object { $_.Id -ieq $id } | Select-Object -First 1
        $copy = if ($def) { Resolve-MswpoList $def } else { $null }
        $t = [pscustomobject]@{ List = $copy; Broken = $true; Def = $def }
    }
    $LookupTargetCache[$id] = $t
    return $t
}

function Resolve-LookupId($Field, [string]$Text) {
    $t = Find-LookupTarget $Field.LookupList
    if (-not $t -or -not $t.List) {
        throw "Lookup column '$($Field.Title)' points to list $($Field.LookupList), which does not exist on this site and is not a known MSWPO list. Run -Action Check -Step Lookups."
    }
    if ($t.Broken) {
        Write-Warning "Lookup column '$($Field.Title)' still points to the PRODUCTION list '$($t.Def.Title)' (id $($Field.LookupList)), which does not exist on this site. Using the test copy for the item id - but the app cannot read this column until it is repaired: -Action FixLookups."
    }
    $match = Get-Items $t.List | Where-Object { (Format-Value $_.FieldValues[$Field.ShowField]) -eq $Text } | Select-Object -Last 1
    if (-not $match) { throw "No item '$Text' in '$($t.List.Title)' for lookup '$($Field.Title)'" }
    return [int]$match.Id
}

# Every lookup column of the MSWPO lists whose target list does not exist on this site.
function Get-BrokenLookups {
    $out = @()
    foreach ($def in @($Lists.Core.Values) + @($Lists.Station)) {
        $list = Resolve-MswpoList $def
        foreach ($f in ((Get-FieldInfo $list).Values | Where-Object { $_.Type -like 'Lookup*' -and $_.LookupList } | Sort-Object Name)) {
            $t = Find-LookupTarget $f.LookupList
            if ($t -and $t.Broken) { $out += [pscustomobject]@{ List = $list; Field = $f; Target = $t } }
        }
    }
    return $out
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
if ($DeviceLogin) {
    $signIn = @{ Url = $SiteUrl; DeviceLogin = $true; ClientId = $ClientId }
    if ($Tenant) { $signIn.Tenant = $Tenant }
    Connect-PnPOnline @signIn
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
}
Write-Host "Data folder: $DataDir"
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

# The account whose role is changed: -User <email> (e.g. the account signed in to the browser), default = you.
function Get-TargetUser {
    if (-not $User) { return $Me }
    $u = New-PnPUser -LoginName $User   # ensures the user exists on the site and returns it (Id, Title, Email)
    return $u
}

function Get-UserKey($Who) { return ([string]$Who.Email).ToLowerInvariant() }

# Old state files kept one list of item ids for the signed-in account; move it to the per-user map.
function Get-OriginalRoleMap($St) {
    if (-not $St.ContainsKey('originalRolesByUser')) { $St.originalRolesByUser = @{} }
    if ($St.ContainsKey('originalRoles')) {
        $St.originalRolesByUser[(Get-UserKey $Me)] = @($St.originalRoles)
        $St.Remove('originalRoles')
    }
    return $St.originalRolesByUser
}

function Get-ExpectedRole($Items, $Who = $Me) {
    # Legacy rule (roleLogic.ts): ID order, last match wins, Supervisors is never replaced.
    $resolved = ''
    foreach ($it in ($Items | Sort-Object { [int]$_.Id })) {
        if ((Get-MemberIds $it) -contains $Who.Id -and $resolved -ne 'Supervisors') { $resolved = [string]$it.FieldValues['Title'] }
    }
    $isOperator = (Find-OperatorItems (Get-Core 'operatorLookup') 'Operator_x0020_Name' $Who.Title).Count -gt 0
    if ($resolved -eq 'Trainer') { if ($isOperator) { return 'Trainer-Operator' } else { return 'Trainer' } }
    if (-not $resolved) { if ($isOperator) { return 'Operator' } else { return 'Visitor' } }
    return $resolved
}

function Set-UserGroups($Items, $Who, [int[]]$WantIds) {
    foreach ($it in $Items) {
        $ids = Get-MemberIds $it
        $isIn = $ids -contains $Who.Id
        $want = $WantIds -contains [int]$it.Id
        if ($want -and -not $isIn) { Set-Membership $it ($ids + $Who.Id); Write-Host "  added to     $($it.FieldValues['Title']) (#$($it.Id))" }
        if (-not $want -and $isIn) { Set-Membership $it ($ids | Where-Object { $_ -ne $Who.Id }); Write-Host "  removed from $($it.FieldValues['Title']) (#$($it.Id))" }
    }
}

function Invoke-Role([string]$Target) {
    $who = Get-TargetUser
    $list = Get-Core 'accessManagement'
    $items = Get-Items $list | Sort-Object { [int]$_.Id }
    $st = $State
    if (-not $st) { $st = @{ created = @() } }
    $map = Get-OriginalRoleMap $st
    $key = Get-UserKey $who
    if (-not $map.ContainsKey($key)) {
        $map[$key] = @($items | Where-Object { (Get-MemberIds $_) -contains $who.Id } | ForEach-Object { [int]$_.Id })
        Write-Host "Saved the original groups of $($who.Title) (item ids: $(if ($map[$key].Count) { $map[$key] -join ', ' } else { 'none' }))."
    }
    Save-State $st
    $group = if ($Target -eq 'Operator') { '' } else { $Target }
    $targetItem = $items | Where-Object { [string]$_.FieldValues['Title'] -eq $group } | Select-Object -Last 1
    if ($group -and -not $targetItem) { throw "No Access Management item titled '$group'." }
    Write-Host "Setting the role of $($who.Title) <$($who.Email)> to $Target" -ForegroundColor Cyan
    Set-UserGroups $items $who $(if ($targetItem) { @([int]$targetItem.Id) } else { @() })
    $role = Get-ExpectedRole (Get-Items $list) $who
    Write-Host "Role the app should now show for $($who.Title): $role  (reload the page)" -ForegroundColor Green
    return $st
}

# Restores every account whose role this run changed (or only -User).
function Invoke-RestoreRole {
    if (-not $State) { Write-Host 'Nothing to restore (Role was never used in this run).'; return }
    $map = Get-OriginalRoleMap $State
    $keys = @($map.Keys)
    if ($User) { $keys = @($keys | Where-Object { $_ -eq $User.ToLowerInvariant() }) }
    if (-not $keys.Count) { Write-Host 'Nothing to restore (Role was never used in this run).'; return }
    $list = Get-Core 'accessManagement'
    foreach ($k in $keys) {
        $who = New-PnPUser -LoginName $k
        Write-Host "Restoring the groups of $($who.Title) <$k>" -ForegroundColor Cyan
        Set-UserGroups (Get-Items $list) $who @($map[$k])
        Write-Host "  role the app should show: $(Get-ExpectedRole (Get-Items $list) $who)" -ForegroundColor Green
        $map.Remove($k)
    }
    Save-State $State
}

# ---- setup (section A) ----------------------------------------------------------------------------
function Get-CreatedCount($St, [string]$Key) { return @($St.created | Where-Object { $_.list -eq $Key }).Count }

# Re-running Setup resumes: items already created by this run are skipped, the rest is created.
function Invoke-Setup {
    $lookup = Get-Core 'operatorLookup'
    $lf = Get-FieldInfo $lookup
    if ($State -and $State.operator -and $State.baseline) {
        if ($OperatorName -and $OperatorName -ne $State.operator) { throw "A run for '$($State.operator)' already exists. Run -Action Cleanup -Force first." }
        $st = $State
        $name = $st.operator
        $email = $st.email
        Write-Host "Resuming setup for '$name' (already created: $(($st.created | ForEach-Object { $_.list }) -join ', '))" -ForegroundColor Yellow
    } else {
        $name = if ($OperatorName) { $OperatorName } else { $Me.Title }
        $email = if ($OperatorEmail) { $OperatorEmail } else { $Me.Email }
        if ((Find-OperatorItems $lookup 'Operator_x0020_Name' $name).Count) { throw "'$name' is already in the Operator LookUp List on this site - pick another -OperatorName or clean up first." }
        $st = if ($State) { $State } else { @{} }
        $st.operator = $name
        $st.email = $email
        $st.startedAt = (Get-Date).ToString('o')
        $st.created = @()
        Write-Host "Recording the highest item id of every MSWPO list ..."
        $st.baseline = @{}
        foreach ($key in $Lists.Core.Keys) { $st.baseline[$key] = Get-MaxId (Get-Core $key) }
        foreach ($def in $Lists.Station) { $st.baseline[$def.Key] = Get-MaxId (Resolve-MswpoList $def) }
        Save-State $st
    }
    $script:OperatorName = $name
    Write-Host "Test operator '$name' <$email>, shift $Shift" -ForegroundColor Cyan

    # A1 Operator LookUp
    if (-not (Get-CreatedCount $st 'operatorLookup')) {
        $a1 = [ordered]@{
            Operator_x0020_Name = $name; Operator_x0020_Email = $email; Operator_x0020_Shift = $Shift; Job_x0020_Title = 'Operator'
        }
        $a1.Status = Select-ActiveStatus $lf['Status']
        'PrimaryStationHours', 'PrimaryStationTrainingHours', 'SecondaryStationHours', 'SecondaryStationTrainingHours', 'AWTStationHours',
        'AWTStationTrainingHours', 'DWBStationHours', 'DWBStationTrainingHours', 'SupervisorDutiesStationHours',
        'SupervisorDutiesStationTrainingHours', 'SupervisorDutiesHours' | Where-Object { $lf[$_] } | ForEach-Object { $a1[$_] = 0 }
        Add-TrackedItem $st 'operatorLookup' $lookup (New-Values $lookup $a1) | Out-Null
    }

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
    if (-not (Get-CreatedCount $st 'operatorCertification')) {
        Add-TrackedItem $st 'operatorCertification' $emp (New-Values $emp $a2) | Out-Null
    }

    # A3 Safety, last year and this year (note: the comma binds tighter than "-", hence the parentheses)
    $safety = Get-Core 'safetyTraining'
    $sf = Get-FieldInfo $safety
    $year = (Get-Date).Year
    $safetyRows = @(
        @{ Year = ($year - 1); Values = $SafetyLastYear },
        @{ Year = $year; Values = $SafetyThisYear }
    )
    $done = Get-CreatedCount $st 'safetyTraining'
    for ($i = $done; $i -lt $safetyRows.Count; $i++) {
        $row = $safetyRows[$i]
        $y = if ($sf['Year'] -and $sf['Year'].Type -match 'Number') { [int]$row.Year } else { [string]$row.Year }
        $a3 = [ordered]@{
            Operator_x0020_Look_x0020_up = $name; Year = $y
            _x0025_Face_x0020_to_x0020_Face_ = $row.Values[0]; _x0025_Online_x0020_Completed = $row.Values[1]; _x0025__x0020_Total_x0020_Safety = $row.Values[2]
        }
        Add-TrackedItem $st 'safetyTraining' $safety (New-Values $safety $a3) | Out-Null
    }

    # A4 Lab DOC (dates blank)
    $doc = Get-Core 'docCertificationDates'
    $df = Get-FieldInfo $doc
    $a4 = [ordered]@{ Title = $name; Shift = $Shift }
    if ($df['Status']) { $a4.Status = Select-ActiveStatus $df['Status'] }
    if (-not (Get-CreatedCount $st 'docCertificationDates')) {
        Add-TrackedItem $st 'docCertificationDates' $doc (New-Values $doc $a4) | Out-Null
    }

    # A5 Document set (name without a trailing dot - legacy GetDocumentSetURL), Title = operator name
    if (Get-CreatedCount $st 'trainingDocumentSets') { Write-Host 'Setup complete.' -ForegroundColor Green; return $st }
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
        # Noise: list items' file name "64_.000", and legacy calculated columns that hold <img onload=...> script.
        if ($f.Name -eq 'FileLeafRef' -and $v -match '^\d+_\.000$') { continue }
        if ($v -like '<img *') { continue }
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
            $l = @(Find-OperatorItems (Get-Core 'operatorLookup') 'Operator_x0020_Name' $name)
            Write-Check ($l.Count -eq 1) "Operator LookUp: exactly one item ($($l.Count))"
            if ($l.Count) { Write-Check ((Format-Value $l[0].FieldValues['Status']) -ne 'Inactive') "Operator LookUp: status '$(Format-Value $l[0].FieldValues['Status'])' is not Inactive" }
            $e = @(Find-OperatorItems (Get-Core 'operatorCertification') 'System_x0020_Defined_x0020_Opera' $name)
            Write-Check ($e.Count -eq 1) "Employee Data: exactly one item ($($e.Count))"
            if ($e.Count) {
                foreach ($f in 'Primary_x0020_Current_x0020_Leve', 'Secondary_x0020_Current_x0020_Le', 'AWT_x0020_Current_x0020_Level', 'DWB_x0020_Current_x0020_Level') {
                    Write-Check ((Format-Value $e[0].FieldValues[$f]) -eq 'T') "Employee Data: $f = T ($(Format-Value $e[0].FieldValues[$f]))"
                }
                Write-Check ((Format-Value $e[0].FieldValues['Shift_x0020_Supervisor']) -ne '') "Employee Data: shift supervisor set ($(Format-Value $e[0].FieldValues['Shift_x0020_Supervisor']))"
            }
            $s = @(Find-OperatorItems (Get-Core 'safetyTraining') 'Operator_x0020_Look_x0020_up' $name)
            Write-Check ($s.Count -ge 1) "Safety: $($s.Count) year(s): $(($s | ForEach-Object { Format-Value $_.FieldValues['Year'] }) -join ', ')"
            $d = @(Find-OperatorItems (Get-Core 'docCertificationDates') 'Title' $name)
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
        'Lookups' {
            $broken = Get-BrokenLookups
            foreach ($b in $broken) {
                $copy = if ($b.Target.List) { "test copy: '$($b.Target.List.Title)'" } else { 'no copy found' }
                $was = if ($b.Target.Def) { "production '$($b.Target.Def.Title)'" } else { 'an unknown list' }
                Write-Check $false "$($b.List.Title) > $($b.Field.Title) ($($b.Field.Name)) points to $was $($b.Field.LookupList) - $copy"
            }
            if (-not $broken.Count) { Write-Check $true 'Every lookup column of the MSWPO lists points to a list on this site' }
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

# Re-points lookup columns that still reference production lists to the test copies (test site schema only).
function Invoke-FixLookups {
    $broken = @(Get-BrokenLookups)
    if (-not $broken.Count) { Write-Host 'No broken lookup columns.' -ForegroundColor Green; return }
    $broken | ForEach-Object {
        [pscustomobject]@{
            List = $_.List.Title; Column = $_.Field.Title; 'Points to (missing)' = $_.Field.LookupList
            'Will point to' = if ($_.Target.List) { "$($_.Target.List.Title) {$($_.Target.List.Id)}" } else { '(no copy - cannot fix)' }
        }
    } | Format-Table -AutoSize -Wrap | Out-String | Write-Host
    if (-not $Force) { Write-Host 'Dry run. Add -Force to change these columns on the TEST site.' -ForegroundColor Yellow; return }
    $webId = (Get-PnPWeb).Id
    foreach ($b in $broken) {
        if (-not $b.Target.List) { Write-Warning "Skipped $($b.List.Title) > $($b.Field.Title): no matching list on this site."; continue }
        try {
            $field = Get-PnPField -List $b.List -Identity $b.Field.Name
            $xml = $field.SchemaXml.Replace("List=`"$($b.Field.LookupList)`"", "List=`"{$($b.Target.List.Id)}`"")
            $xml = [regex]::Replace($xml, 'WebId="[^"]*"', "WebId=`"$webId`"")
            $field.SchemaXml = $xml
            $field.Update()
            Invoke-PnPQuery
            Write-Host "  fixed $($b.List.Title) > $($b.Field.Title)" -ForegroundColor Green
        } catch {
            Write-Warning "  $($b.List.Title) > $($b.Field.Title): SharePoint refused the change ($($_.Exception.Message)). The column would have to be recreated."
        }
    }
    $FieldCache.Clear()
    $LookupTargetCache.Clear()
    $left = @(Get-BrokenLookups)
    Write-Host "Broken lookup columns left: $($left.Count)"
    Write-Host 'Now compare an existing operator (Admin page safety table, profile safety chart) with production: copied items keep their old lookup ids.' -ForegroundColor Yellow
}

function Invoke-Status {
    if (-not $State) { Write-Host 'No run set up yet (use -Action Setup).'; return }
    Write-Host "Operator: $($State.operator) <$($State.email)>  started $($State.startedAt)"
    Write-Host "Created by Setup: $(($State.created | ForEach-Object { "$($_.list)#$($_.id)" }) -join ', ')"
    $items = Get-Items (Get-Core 'accessManagement')
    Write-Host "Your role (signed-in account): $(Get-ExpectedRole $items)"
    $map = Get-OriginalRoleMap $State
    foreach ($k in $map.Keys) {
        $who = New-PnPUser -LoginName $k
        Write-Host "Changed by this run: $($who.Title) <$k> - now $(Get-ExpectedRole $items $who); original groups (item ids): $(if (@($map[$k]).Count) { @($map[$k]) -join ', ' } else { 'none' })"
    }
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
    'FixLookups' { Invoke-FixLookups }
}
