[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TargetUser,

    [Parameter(Mandatory)]
    [string]$SpoAdminUrl,

    [pscredential]$Credential,

    [switch]$InteractiveAuth,

    [string]$AdminLogin,

    [switch]$Execute,

    [switch]$NoExecutePrompt,

    [string]$OutputRoot,

    [int]$ProgressEvery = 25,

    [string]$OnlySitesCsv,

    [switch]$OnlyPersonalSites,

    [switch]$SkipPersonalSites
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$onlyPersonalSitesMode = [bool]$OnlyPersonalSites.IsPresent
$skipPersonalSitesMode = [bool]$SkipPersonalSites.IsPresent

if ($onlyPersonalSitesMode -and $skipPersonalSitesMode) {
    throw "Use only one site scope switch: -OnlyPersonalSites or -SkipPersonalSites."
}

$scriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $PWD.Path
}

function ConvertTo-SafeName {
    param([string]$Value)
    return ($Value -replace "^https?://", "" -replace "[^a-zA-Z0-9._-]+", "_").Trim("_")
}

function Resolve-SpoAdminUrl {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "SPO admin URL could not be resolved. Provide a tenant name, SharePoint URL, OneDrive URL, or admin URL."
    }

    $candidate = $Value.Trim()

    if ($candidate -match "^[a-zA-Z0-9-]+$") {
        return "https://$candidate-admin.sharepoint.com"
    }

    if ($candidate -notmatch "^https?://" -and $candidate -match "^[^/]+\.sharepoint\.com") {
        $candidate = "https://$candidate"
    }

    $uri = $null
    if (-not [System.Uri]::TryCreate($candidate, [System.UriKind]::Absolute, [ref]$uri)) {
        throw "SPO admin URL could not be resolved from '$Value'. Use a tenant name or a SharePoint URL."
    }

    $hostName = $uri.Host.ToLowerInvariant()
    if ($hostName -notmatch "^(.+?)(-admin|-my)?\.sharepoint\.com$") {
        throw "SPO admin URL could not be resolved from '$Value'. Host '$hostName' is not a SharePoint Online host."
    }

    $tenantName = $matches[1]
    return "https://$tenantName-admin.sharepoint.com"
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $reportUserFolder = ConvertTo-SafeName $TargetUser
    $OutputRoot = Join-Path (Join-Path (Join-Path $scriptRoot "Reports") $reportUserFolder) ("Run-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$null = New-Item -Path $OutputRoot -ItemType Directory -Force
$exportRoot = Join-Path $OutputRoot "exports"
$null = New-Item -Path $exportRoot -ItemType Directory -Force
$executeRemediation = [bool]$Execute.IsPresent

$allRowsPath = Join-Path $OutputRoot "all-userinfo-rows.csv"
$actionPlanPath = Join-Path $OutputRoot "action-plan.csv"
$resultsPath = Join-Path $OutputRoot "remediation-results.csv"
$errorsPath = Join-Path $OutputRoot "errors.csv"
$summaryPath = Join-Path $OutputRoot "summary.txt"
$commandsPath = Join-Path $OutputRoot "commands-that-would-run.ps1"

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -Path $summaryPath -Value $line -Encoding UTF8
}

function Export-CsvAppend {
    param(
        [object[]]$Rows,
        [Parameter(Mandatory)][string]$Path
    )
    if ($null -eq $Rows -or $Rows.Count -eq 0) { return }
    if (Test-Path $Path) {
        $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Append
    } else {
        $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
}

function Get-ExportFileName {
    param([string]$User)
    return (($User -replace "[^a-zA-Z0-9]+", "_").Trim("_") + "_userinfo.csv")
}

function Get-SpoExportCsv {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$Suffix,
        [datetime]$After
    )

    $files = @(Get-ChildItem -Path $Folder -Filter "*$Suffix" -File -ErrorAction SilentlyContinue |
        Where-Object { $null -eq $After -or $_.LastWriteTime -ge $After } |
        Sort-Object LastWriteTime -Descending)

    if ($files.Count -eq 0) { return $null }
    return $files[0].FullName
}

function Export-UserInfoRows {
    param(
        [string]$SiteUrl,
        [string]$User,
        [string]$Folder
    )
    $null = New-Item -Path $Folder -ItemType Directory -Force
    $exportStarted = Get-Date
    Export-SPOUserInfo -Site $SiteUrl -LoginName $User -OutputFolder $Folder -ErrorAction Stop | Out-Null
    $file = Join-Path $Folder (Get-ExportFileName -User $User)
    if (-not (Test-Path $file)) {
        $file = Get-SpoExportCsv -Folder $Folder -Suffix "_userinfo.csv" -After $exportStarted
    }
    if ([string]::IsNullOrWhiteSpace($file) -or -not (Test-Path $file)) { return @() }
    return @(Import-Csv $file)
}

function Get-SiteState {
    param(
        [object[]]$Rows,
        [string]$CurrentSystemUserKey
    )
    if ($Rows.Count -eq 0) { return "NotPresent" }
    $hasCurrent = [bool]($Rows | Where-Object { $_.SystemUserKey -eq $CurrentSystemUserKey })
    $hasStale = [bool]($Rows | Where-Object { $_.SystemUserKey -and $_.SystemUserKey -ne $CurrentSystemUserKey })
    if ($hasCurrent -and $hasStale) { return "OldAndCurrent" }
    if ($hasCurrent) { return "OnlyCurrent" }
    if ($hasStale) { return "OnlyStale" }
    return "Unknown"
}

function Test-PersonalSiteUrl {
    param([string]$Url)
    return ($Url -like "*-my.sharepoint.com/personal/*")
}

function Get-ScopedSites {
    param(
        [string]$SitesCsv,
        [bool]$OnlyPersonal,
        [bool]$SkipPersonal
    )

    if (-not [string]::IsNullOrWhiteSpace($SitesCsv)) {
        if (-not (Test-Path $SitesCsv)) {
            throw "OnlySitesCsv file not found: $SitesCsv"
        }

        $siteUrls = @(Import-Csv -Path $SitesCsv | ForEach-Object {
            if ($_.PSObject.Properties.Name -contains "SiteUrl") {
                $_.SiteUrl
            } elseif ($_.PSObject.Properties.Name -contains "Url") {
                $_.Url
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

        if ($siteUrls.Count -eq 0) {
            throw "OnlySitesCsv must contain at least one non-empty SiteUrl or Url value."
        }

        if ($OnlyPersonal) {
            $siteUrls = @($siteUrls | Where-Object { Test-PersonalSiteUrl -Url $_ })
        } elseif ($SkipPersonal) {
            $siteUrls = @($siteUrls | Where-Object { -not (Test-PersonalSiteUrl -Url $_) })
        }

        $sites = foreach ($siteUrl in $siteUrls) {
            try {
                Get-SPOSite -Identity $siteUrl -ErrorAction Stop
            } catch {
                Write-Log "Failed to load scoped site '$siteUrl': $($_.Exception.Message)"
            }
        }

        return @($sites)
    }

    if ($OnlyPersonal) {
        try {
            return @(Get-SPOSite -Limit All -IncludePersonalSite $true -Filter "Url -like '-my.sharepoint.com/personal/'" -ErrorAction Stop)
        } catch {
            Write-Log "Server-side personal site filter failed; falling back to local filtering. Error: $($_.Exception.Message)"
            return @(Get-SPOSite -Limit All -IncludePersonalSite $true | Where-Object { Test-PersonalSiteUrl -Url $_.Url })
        }
    }

    if ($SkipPersonal) {
        return @(Get-SPOSite -Limit All -IncludePersonalSite $false)
    }

    return @(Get-SPOSite -Limit All -IncludePersonalSite $true)
}

function Invoke-TargetRemoval {
    param(
        [string]$SiteUrl,
        [string]$User,
        [string]$AdminLogin
    )

    $claimsLogin = "i:0#.f|membership|$User"
    $attempts = @()
    $adminAdded = $false
    $adminWasSiteAdmin = $false
    $checkedAdmin = $false

    try {
        try {
            $adminUser = Get-SPOUser -Site $SiteUrl -LoginName $AdminLogin -ErrorAction Stop
            $checkedAdmin = $true
            $adminWasSiteAdmin = [bool]$adminUser.IsSiteAdmin
        } catch {
            $checkedAdmin = $false
        }

        try {
            Remove-SPOUser -Site $SiteUrl -LoginName $claimsLogin -ErrorAction Stop | Out-Null
            $attempts += "Remove-SPOUser claims: success"
        } catch {
            $attempts += "Remove-SPOUser claims before admin: $($_.Exception.Message)"
            Set-SPOUser -Site $SiteUrl -LoginName $AdminLogin -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
            $adminAdded = $true
            Remove-SPOUser -Site $SiteUrl -LoginName $claimsLogin -ErrorAction Stop | Out-Null
            $attempts += "Remove-SPOUser claims after temp admin: success"
        }

        try {
            Remove-SPOUser -Site $SiteUrl -LoginName $User -ErrorAction Stop | Out-Null
            $attempts += "Remove-SPOUser simple: success"
        } catch {
            $attempts += "Remove-SPOUser simple: $($_.Exception.Message)"
        }

        return [pscustomobject]@{
            Success          = $true
            TempAdminAdded   = $adminAdded
            AdminWasSiteAdmin = $adminWasSiteAdmin
            CheckedAdmin     = $checkedAdmin
            Attempts         = ($attempts -join " | ")
        }
    } catch {
        return [pscustomobject]@{
            Success          = $false
            TempAdminAdded   = $adminAdded
            AdminWasSiteAdmin = $adminWasSiteAdmin
            CheckedAdmin     = $checkedAdmin
            Attempts         = (($attempts + "Fatal: $($_.Exception.Message)") -join " | ")
        }
    } finally {
        if ($adminAdded -and -not $adminWasSiteAdmin) {
            try {
                Set-SPOUser -Site $SiteUrl -LoginName $AdminLogin -IsSiteCollectionAdmin $false -ErrorAction Stop | Out-Null
            } catch {
                Export-CsvAppend -Rows @([pscustomobject]@{
                    SiteUrl = $SiteUrl
                    Stage   = "RemoveTempAdmin"
                    Error   = $_.Exception.Message
                }) -Path $errorsPath
            }
        }
    }
}

function Invoke-PlannedSiteRemediation {
    param(
        [pscustomobject]$PlannedSite,
        [string]$User,
        [string]$AdminLogin,
        [string]$CurrentSystemUserKey,
        [object]$Counts,
        [System.Collections.Generic.List[string]]$RemediatedSites
    )

    $removal = Invoke-TargetRemoval -SiteUrl $PlannedSite.SiteUrl -User $User -AdminLogin $AdminLogin
    if ($removal.Success) {
        $Counts.Remediated++
        $RemediatedSites.Add($PlannedSite.SiteUrl) | Out-Null
    } else {
        $Counts.RemediationFailed++
    }

    $afterRows = @()
    try {
        $afterRows = @(Export-UserInfoRows -SiteUrl $PlannedSite.SiteUrl -User $User -Folder (Join-Path $PlannedSite.SiteExportFolder "after"))
    } catch {
        Export-CsvAppend -Rows @([pscustomobject]@{
            SiteUrl = $PlannedSite.SiteUrl
            Stage   = "AfterExport"
            Error   = $_.Exception.Message
        }) -Path $errorsPath
    }

    Export-CsvAppend -Rows @([pscustomobject]@{
        SiteUrl           = $PlannedSite.SiteUrl
        SiteTitle         = $PlannedSite.SiteTitle
        SiteOwner         = $PlannedSite.SiteOwner
        BeforeState       = $PlannedSite.State
        AfterState        = Get-SiteState -Rows $afterRows -CurrentSystemUserKey $CurrentSystemUserKey
        Success           = $removal.Success
        TempAdminAdded    = $removal.TempAdminAdded
        AdminWasSiteAdmin = $removal.AdminWasSiteAdmin
        Attempts          = $removal.Attempts
        Note              = "If AfterState remains OnlyStale, ask owner to share again; SharePoint should recreate the current SystemUserKey."
    }) -Path $resultsPath
}

Import-Module Microsoft.Online.SharePoint.PowerShell

$resolvedSpoAdminUrl = Resolve-SpoAdminUrl -Value $SpoAdminUrl

Write-Log "Mode: $(if ($executeRemediation) { 'EXECUTE' } else { 'DRY-RUN' })"
if ($resolvedSpoAdminUrl -ne $SpoAdminUrl) {
    Write-Log "Resolved SPO admin URL: $resolvedSpoAdminUrl"
}
if ($InteractiveAuth -or $null -eq $Credential) {
    Write-Log "Connecting to SPO admin with interactive authentication: $resolvedSpoAdminUrl"
    Connect-SPOService -Url $resolvedSpoAdminUrl
} else {
    Write-Log "Connecting to SPO admin with credential authentication: $resolvedSpoAdminUrl"
    Connect-SPOService -Url $resolvedSpoAdminUrl -Credential $Credential
}

Write-Log "Exporting current User Profile for $TargetUser"
$profileExportStarted = Get-Date
Export-SPOUserProfile -LoginName $TargetUser -OutputFolder $OutputRoot
$profileFile = Join-Path $OutputRoot (($TargetUser -replace "[^a-zA-Z0-9]+", "_").Trim("_") + "_profile.csv")
if (-not (Test-Path $profileFile)) {
    $profileFile = Get-SpoExportCsv -Folder $OutputRoot -Suffix "_profile.csv" -After $profileExportStarted
}
if ([string]::IsNullOrWhiteSpace($profileFile) -or -not (Test-Path $profileFile)) {
    throw "Profile export not found in output folder: $OutputRoot"
}
Write-Log "Profile export file: $profileFile"

$profile = Import-Csv $profileFile | Select-Object -First 1
$currentSystemUserKey = $profile.SID
$currentObjectId = $profile."msOnline-ObjectId"
$currentLoginName = $profile.AccountName

if ([string]::IsNullOrWhiteSpace($currentSystemUserKey)) {
    throw "Current SystemUserKey/SID not found in profile export."
}

Write-Log "Current LoginName: $currentLoginName"
Write-Log "Current SystemUserKey: $currentSystemUserKey"
Write-Log "Current ObjectID: $currentObjectId"

if (-not [string]::IsNullOrWhiteSpace($OnlySitesCsv)) {
    $siteScope = "OnlySitesCsv"
} elseif ($onlyPersonalSitesMode) {
    $siteScope = "OnlyPersonalSites"
} elseif ($skipPersonalSitesMode) {
    $siteScope = "SkipPersonalSites"
} else {
    $siteScope = "AllSites"
}

Write-Log "Enumerating site collections. Scope: $siteScope"
$sites = @(Get-ScopedSites -SitesCsv $OnlySitesCsv -OnlyPersonal $onlyPersonalSitesMode -SkipPersonal $skipPersonalSitesMode)
Write-Log "Total site collections: $($sites.Count)"

"# Generated $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" | Set-Content -Path $commandsPath -Encoding UTF8
"# Execute mode in this run: $executeRemediation" | Add-Content -Path $commandsPath -Encoding UTF8
"# Site scope: $siteScope" | Add-Content -Path $commandsPath -Encoding UTF8
if (-not [string]::IsNullOrWhiteSpace($OnlySitesCsv)) {
    "# OnlySitesCsv: $OnlySitesCsv" | Add-Content -Path $commandsPath -Encoding UTF8
}
"# Sites with OldAndCurrent are skipped deliberately." | Add-Content -Path $commandsPath -Encoding UTF8

$counts = [ordered]@{
    SitesScanned      = 0
    NotPresent        = 0
    OnlyCurrent       = 0
    OldAndCurrent     = 0
    OnlyStale         = 0
    Unknown           = 0
    Remediated        = 0
    RemediationFailed = 0
    Errors            = 0
}

if (-not [string]::IsNullOrWhiteSpace($AdminLogin)) {
    $adminLogin = $AdminLogin
} elseif ($null -ne $Credential) {
    $adminLogin = $Credential.UserName
} else {
    $adminLogin = $null
}

if ($executeRemediation -and [string]::IsNullOrWhiteSpace($adminLogin)) {
    throw "Execute mode with interactive authentication requires -AdminLogin so temporary Site Collection Admin assignment can be removed correctly."
}

$remediatedSites = New-Object System.Collections.Generic.List[string]
$plannedOnlyStaleSites = New-Object System.Collections.Generic.List[object]
$promptedRemediationExecuted = $false

foreach ($site in $sites) {
    $counts.SitesScanned++
    $siteUrl = $site.Url
    if ($counts.SitesScanned % $ProgressEvery -eq 0) {
        Write-Log "Progress: $($counts.SitesScanned) / $($sites.Count)"
    }

    $safeName = ConvertTo-SafeName $siteUrl
    $siteExportFolder = Join-Path $exportRoot $safeName

    try {
        $beforeRows = @(Export-UserInfoRows -SiteUrl $siteUrl -User $TargetUser -Folder (Join-Path $siteExportFolder "before"))
        $state = Get-SiteState -Rows $beforeRows -CurrentSystemUserKey $currentSystemUserKey
        $counts[$state]++

        foreach ($row in $beforeRows) {
            Export-CsvAppend -Rows @([pscustomobject]@{
                SiteUrl              = $row.SiteUrl
                SiteTitle            = $site.Title
                SiteOwner            = $site.Owner
                SiteTemplate         = $site.Template
                LockState            = $site.LockState
                LoginName            = $row.LoginName
                Email                = $row.Email
                Title                = $row.Title
                SystemUserKey        = $row.SystemUserKey
                CurrentSystemUserKey = $currentSystemUserKey
                CurrentObjectId      = $currentObjectId
                State                = $state
            }) -Path $allRowsPath
        }

        $plan = [pscustomobject]@{
            SiteUrl              = $siteUrl
            SiteTitle            = $site.Title
            SiteOwner            = $site.Owner
            SiteTemplate         = $site.Template
            LockState            = $site.LockState
            State                = $state
            Action               = if ($state -eq "OnlyStale") { "RemoveSPOUserAndRequireReshare" } else { "Skip" }
            Reason               = switch ($state) {
                "OnlyStale"     { "Only legacy SystemUserKey exists; remove access principal so a new share can recreate current identity." }
                "OldAndCurrent" { "Legacy and current entries both exist; user should already resolve with current identity. Do not touch." }
                "OnlyCurrent"   { "Only current identity exists." }
                "NotPresent"    { "Target user not present in this site." }
                default         { "Unclassified state." }
            }
            CurrentSystemUserKey = $currentSystemUserKey
            FoundSystemUserKeys  = (($beforeRows | ForEach-Object SystemUserKey | Sort-Object -Unique) -join "; ")
        }
        Export-CsvAppend -Rows @($plan) -Path $actionPlanPath

        if ($state -eq "OnlyStale") {
            'Remove-SPOUser -Site "{0}" -LoginName "i:0#.f|membership|{1}"' -f $siteUrl, $TargetUser |
                Add-Content -Path $commandsPath -Encoding UTF8
            'Remove-SPOUser -Site "{0}" -LoginName "{1}"' -f $siteUrl, $TargetUser |
                Add-Content -Path $commandsPath -Encoding UTF8

            $plannedOnlyStaleSites.Add([pscustomobject]@{
                SiteUrl          = $siteUrl
                SiteTitle        = $site.Title
                SiteOwner        = $site.Owner
                SiteExportFolder = $siteExportFolder
                State            = $state
            }) | Out-Null

            if ($executeRemediation) {
                Invoke-PlannedSiteRemediation `
                    -PlannedSite $plannedOnlyStaleSites[$plannedOnlyStaleSites.Count - 1] `
                    -User $TargetUser `
                    -AdminLogin $adminLogin `
                    -CurrentSystemUserKey $currentSystemUserKey `
                    -Counts $counts `
                    -RemediatedSites $remediatedSites
            }
        }
    } catch {
        $counts.Errors++
        Export-CsvAppend -Rows @([pscustomobject]@{
            SiteUrl      = $siteUrl
            SiteTitle    = $site.Title
            SiteOwner    = $site.Owner
            SiteTemplate = $site.Template
            LockState    = $site.LockState
            Stage        = "ScanSite"
            Error        = $_.Exception.Message
        }) -Path $errorsPath
    }
}

if (-not $executeRemediation -and $plannedOnlyStaleSites.Count -gt 0 -and -not $NoExecutePrompt) {
    Write-Host ""
    $answer = Read-Host "Dry-run found $($plannedOnlyStaleSites.Count) OnlyStale site(s). Execute remediation now using these results only? (Y/N)"
    if ($answer -match "^(y|yes|s|sim)$") {
        if ([string]::IsNullOrWhiteSpace($adminLogin)) {
            $adminLogin = Read-Host "Admin login used for temporary Site Collection Admin assignment"
        }

        if ([string]::IsNullOrWhiteSpace($adminLogin)) {
            throw "Admin login is required to execute remediation from dry-run results."
        }

        Write-Log "Executing remediation from dry-run results only: $($plannedOnlyStaleSites.Count) site(s)"
        foreach ($plannedSite in $plannedOnlyStaleSites) {
            Invoke-PlannedSiteRemediation `
                -PlannedSite $plannedSite `
                -User $TargetUser `
                -AdminLogin $adminLogin `
                -CurrentSystemUserKey $currentSystemUserKey `
                -Counts $counts `
                -RemediatedSites $remediatedSites
        }
        $promptedRemediationExecuted = $true
    } else {
        Write-Log "Remediation prompt skipped by operator."
    }
}

Write-Log "Completed"
foreach ($key in $counts.Keys) {
    Write-Log "$key`: $($counts[$key])"
}

[pscustomobject]@{
    OutputRoot           = $OutputRoot
    Execute              = $executeRemediation
    PromptedRemediationExecuted = $promptedRemediationExecuted
    SpoAdminUrlInput     = $SpoAdminUrl
    SpoAdminUrl          = $resolvedSpoAdminUrl
    AuthMode             = if ($InteractiveAuth -or $null -eq $Credential) { "Interactive" } else { "Credential" }
    AdminLogin           = $adminLogin
    SiteScope            = $siteScope
    OnlySitesCsv         = $OnlySitesCsv
    CurrentLoginName     = $currentLoginName
    CurrentSystemUserKey = $currentSystemUserKey
    CurrentObjectId      = $currentObjectId
    Counts               = $counts
    RemediatedSites      = @($remediatedSites)
    AllRowsCsv           = $allRowsPath
    ActionPlanCsv        = $actionPlanPath
    ResultsCsv           = $resultsPath
    ErrorsCsv            = $errorsPath
    CommandsFile         = $commandsPath
    Summary              = $summaryPath
} | ConvertTo-Json -Depth 6
