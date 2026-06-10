# M365 Stale SystemUserKey Remediation

PowerShell tool for detecting and remediating stale SharePoint Online
`SystemUserKey` entries for a Microsoft 365 user.

This is useful when a user was recreated, migrated, restored, synchronized
again, or otherwise received a new identity in Microsoft 365, while SharePoint
Online still keeps one or more old user entries in site collections.

The common symptom is that the user exists and appears valid in Microsoft 365,
but access to some SharePoint or OneDrive content still fails, resolves to an
old identity, or behaves inconsistently.

## What the Script Does

`Invoke-StaleSystemUserKeyRemediation.ps1` compares the current SharePoint Online
User Profile `SystemUserKey` for a target user against the user entries found in
SharePoint and OneDrive site collections.

For every site collection, it classifies the user's state:

| State | Meaning | Action |
| --- | --- | --- |
| `NotPresent` | The user is not present in the site collection. | Skip. |
| `OnlyCurrent` | The site only contains the current identity. | Skip. |
| `OldAndCurrent` | The site contains both old and current identities. | Skip deliberately. |
| `OnlyStale` | The site only contains a stale identity. | Plan removal; remove only with `-Execute`. |
| `Unknown` | Rows were present but could not be classified cleanly. | Skip. |

By default, the script runs in dry-run mode. It creates evidence and an action
plan, but it does not remove anything unless `-Execute` is provided.

## Why This Happens

SharePoint Online stores users inside each site collection. These entries can
outlive identity lifecycle events in Microsoft Entra ID, such as:

- user deletion and recreation
- tenant migration
- hybrid identity changes
- immutable ID or synchronization changes
- account restore after deletion
- UPN or mail changes combined with historical sharing entries

When SharePoint keeps a stale `SystemUserKey`, the user may still appear in
permissions, sharing records, or site user information, but the entry no longer
matches the current Microsoft 365 identity.

## Safety Model

The script is intentionally conservative:

- Dry-run is the default.
- Remediation only happens with `-Execute`.
- Only sites classified as `OnlyStale` are remediated.
- Sites classified as `OldAndCurrent` are skipped deliberately.
- Before and after evidence is exported.
- A command file is generated for review.
- If temporary Site Collection Admin elevation is needed, the script attempts to
  restore the previous admin state afterward.

The remediation action is `Remove-SPOUser`, which removes the user entry from
the target site collection. It does not delete the Microsoft 365 user account.

After the stale site user entry is removed, the affected content usually needs to
be shared again or accessed again so SharePoint can recreate the user entry with
the current identity.

## Requirements

- Windows PowerShell 5.1
- SharePoint Online Management Shell module
- SharePoint Administrator or Global Administrator permissions
- Permission to enumerate SharePoint and OneDrive site collections
- Permission to run `Export-SPOUserProfile`, `Export-SPOUserInfo`,
  `Get-SPOSite`, `Get-SPOUser`, `Set-SPOUser`, and `Remove-SPOUser`

Install the SharePoint Online module:

```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

## Files

| File | Purpose |
| --- | --- |
| `Invoke-StaleSystemUserKeyRemediation.ps1` | Main scan and remediation script. |
| `README.md` | Usage and operational guidance. |
| `.gitignore` | Excludes generated reports and local artifacts. |

## Parameters

| Parameter | Required | Description |
| --- | --- | --- |
| `-TargetUser` | Yes | User principal name or login name to inspect. |
| `-SpoAdminUrl` | Yes | SharePoint Online admin URL, such as `https://contoso-admin.sharepoint.com`. |
| `-Credential` | Yes | Credential used for SharePoint Online administration. |
| `-Execute` | No | Applies remediation. Without this switch, the script only creates evidence and an action plan. |
| `-OutputRoot` | No | Custom output folder. Defaults to `Reports/<user>/Run-<timestamp>`. |
| `-ProgressEvery` | No | Progress interval while scanning site collections. Defaults to `25`. |
| `-OnlyPersonalSites` | No | Scans only OneDrive personal sites. Mutually exclusive with `-SkipPersonalSites`. |
| `-SkipPersonalSites` | No | Scans only SharePoint sites, excluding OneDrive personal sites. Mutually exclusive with `-OnlyPersonalSites`. |

## Basic Dry-Run

Start with a dry-run:

```powershell
$cred = Get-Credential

.\Invoke-StaleSystemUserKeyRemediation.ps1 `
  -TargetUser "user@contoso.com" `
  -SpoAdminUrl "https://contoso-admin.sharepoint.com" `
  -Credential $cred
```

This scans SharePoint and OneDrive site collections, classifies each site, and
generates reports. No users are removed.

## Scope the Scan

Scan only OneDrive personal sites:

```powershell
.\Invoke-StaleSystemUserKeyRemediation.ps1 `
  -TargetUser "user@contoso.com" `
  -SpoAdminUrl "https://contoso-admin.sharepoint.com" `
  -Credential $cred `
  -OnlyPersonalSites
```

Scan only SharePoint sites, excluding OneDrive personal sites:

```powershell
.\Invoke-StaleSystemUserKeyRemediation.ps1 `
  -TargetUser "user@contoso.com" `
  -SpoAdminUrl "https://contoso-admin.sharepoint.com" `
  -Credential $cred `
  -SkipPersonalSites
```

For large tenants, scoped runs are easier to review and safer to execute.

## Execute Remediation

After reviewing `action-plan.csv`, run with `-Execute` only when the planned
actions are expected:

```powershell
.\Invoke-StaleSystemUserKeyRemediation.ps1 `
  -TargetUser "user@contoso.com" `
  -SpoAdminUrl "https://contoso-admin.sharepoint.com" `
  -Credential $cred `
  -Execute
```

The script removes the target user only from sites classified as `OnlyStale`.

## Output Structure

By default, output is written to:

```text
Reports/<target-user>/Run-YYYYMMDD-HHMMSS/
```

Generated files:

| File | Description |
| --- | --- |
| `summary.txt` | Timestamped run summary and counters. |
| `all-userinfo-rows.csv` | User information rows found during the scan. |
| `action-plan.csv` | Site-by-site classification and planned action. |
| `commands-that-would-run.ps1` | Reviewable `Remove-SPOUser` commands for `OnlyStale` sites. |
| `remediation-results.csv` | Before/after remediation result, only populated during `-Execute`. |
| `errors.csv` | Site-level scan or remediation errors. |
| `exports/` | Raw SharePoint export data grouped by site collection. |

## Reading the Action Plan

The most important file is `action-plan.csv`.

Recommended interpretation:

| State | Recommended response |
| --- | --- |
| `NotPresent` | No action required. |
| `OnlyCurrent` | No action required. |
| `OldAndCurrent` | Review manually. The current identity already exists, so automatic removal is intentionally skipped. |
| `OnlyStale` | Candidate for `Remove-SPOUser` remediation. |
| `Unknown` | Review manually before taking any action. |

Do not execute remediation based only on the count of stale rows. Review the site
URLs and confirm they match the affected workload or user complaint.

## Temporary Site Collection Admin Behavior

Some sites may not allow `Remove-SPOUser` until the administrative account is a
Site Collection Admin.

When removal fails initially, the script attempts to:

1. Check whether the admin account is already a Site Collection Admin.
2. Temporarily set the admin account as Site Collection Admin.
3. Retry `Remove-SPOUser`.
4. Remove the temporary admin assignment if the account was not originally a
   Site Collection Admin.

Any failure to remove the temporary admin assignment is logged in `errors.csv`.

## Operational Workflow

Recommended process:

1. Run dry-run mode.
2. Review `summary.txt`.
3. Review `action-plan.csv`.
4. Confirm every `OnlyStale` site is expected.
5. Keep a copy of the generated evidence.
6. Run with `-Execute` only after approval.
7. Review `remediation-results.csv`.
8. Ask the owner or source system to share the affected content again if needed.
9. Re-run dry-run mode to confirm the stale entries are gone or replaced by the
   current identity.

## Limitations

- The script depends on SharePoint Online Management Shell behavior.
- Very large tenants can take time to scan.
- `Remove-SPOUser` removes the site collection user entry, not the Entra ID
  account.
- `OldAndCurrent` sites are intentionally skipped because removing the old entry
  automatically may not be necessary and can require business context.
- The script does not inspect every item-level permission directly; it works at
  the site user information level exposed by SharePoint Online exports.

## Troubleshooting

If the profile export fails, confirm:

- the admin URL is correct
- the credential has SharePoint administrator permissions
- the target user exists or can be resolved by SharePoint Online

If personal site filtering fails, the script falls back to local URL filtering.
This is logged in `summary.txt`.

If many sites return errors, verify whether the admin account has permission to
enumerate those site collections.

## Disclaimer

Test in dry-run mode first. Review the exported evidence and action plan before
using `-Execute`. This tool is intended for controlled Microsoft 365 operations
by administrators who understand the impact of removing SharePoint site user
entries.
