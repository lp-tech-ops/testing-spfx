
# O365 PowerShell Scripts – Content-Based Technical Summary

> Scope note: This document summarizes the **scripts explicitly inspected in this session** from your OneDrive folder. Each entry only includes (1) what the script does and (2) the module / API / library usage that is directly visible in the script content. All secrets, IDs, GUIDs, URLs, and account-specific values have been intentionally omitted.

## 1) Site scripts, site designs, and modern SharePoint UX

### AddExtensionSiteScript.ps1
- **What it does:** Creates a SharePoint site script that applies a theme and associates an SPFx Application Customizer extension.
- **Module / platform used:** `Add-SPOSiteScript` from the **SharePoint Online Management Shell**.
- **Notable implementation detail:** The site script JSON includes verbs for theme application and extension association.

### CreateSiteScriptFromExistingSite.ps1
- **What it does:** Extracts a site script from an existing SharePoint site, including branding, theme, regional settings, external sharing capability, and selected lists/libraries; then registers it as a reusable site script.
- **Module / platform used:** `Connect-SPOService`, `Get-SPOSiteScriptFromWeb`, `Add-SPOSiteScript` from the **SharePoint Online Management Shell**.

### CreateSiteDesignFromExistingSite.ps1
- **What it does:** Generates a site script from an existing site and then creates a SharePoint site design based on that site script.
- **Module / platform used:** `Connect-SPOService`, `Get-SPOSiteScriptFromWeb`, `Add-SPOSiteScript`, `Add-SPOSiteDesign` from the **SharePoint Online Management Shell**.

### ConvertListsLibrariesPagesToModern.ps1
- **What it does:** Enables the modern Site Pages feature, converts lists and libraries to modern experience, and enumerates site pages for modernization.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`, `Enable-PnPFeature`, `Get-PnPList`, `Set-PnPList`, `Get-PnPListItem`).

### ChangeListExperience_CPHD.ps1
- **What it does:** Loops through lists/libraries in a SharePoint site and sets each eligible list/library to the **modern** experience while logging results and errors.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`, `Get-PnPList`, `Set-PnPList`).

### ChangeListExperience_DTS_setbacktoclassic.ps1
- **What it does:** Loops through lists/libraries in a SharePoint site and sets each eligible list/library to the **classic** experience while logging results and errors.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`, `Get-PnPList`, `Set-PnPList`).

### ConvertPageLayoutToHome.ps1
- **What it does:** Changes a modern SharePoint page layout to **Home**.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`, `Set-PnPPage`).

### ConvertPageToModern_CPHD.ps1
- **What it does:** Iterates SharePoint site pages, converts classic pages to modern client-side pages, disables comments on the migrated pages, and writes a migration report.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`, `Get-PnPListItem`, `ConvertTo-PnPClientSidePage`, `Set-PnPClientSidePage`).

### ConvertPageToModern_DTS.ps1
- **What it does:** Converts SharePoint site pages to modern pages, locates the migrated page, disables comments, and writes a migration report.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`, `Get-PnPListItem`, `ConvertTo-PnPPage`, `Set-PnPPage`).

### ConvertPageToModern_DTS-Publishing.ps1
- **What it does:** Script scaffold intended for publishing-page modernization work.
- **Module / platform used:** No operational module calls are visible in the inspected portion; only parameter scaffolding and metadata are present.

### ConvertPublishingPagesToModern.ps1
- **What it does:** Enables the modern Site Pages feature as part of a publishing-page modernization workflow.
- **Module / platform used:** **PnP PowerShell** (`Enable-PnPFeature`).

## 2) SharePoint lists, fields, libraries, and file operations

### AddGeoLocationColumn.ps1
- **What it does:** Connects to SharePoint Online, adds a **Geolocation** field to a list, and sets the site property bag value required for map support.
- **Module / platform used:** **SharePoint CSOM** (`Microsoft.SharePoint.Client`, `Microsoft.SharePoint.Client.Runtime`).

### AddGeoLocationColumnusingPnP.ps1
- **What it does:** Adds a **GeoLocation** field to a list and sets a property bag value after connecting interactively.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`, `Get-PnPList`, `Add-PnPField`, `Set-PnPPropertyBagValue`).

### AddIndexColumn.ps1
- **What it does:** Opens a SharePoint list field and marks it as **indexed**.
- **Module / platform used:** **SharePoint CSOM** via `Microsoft.SharePoint.Client` assemblies.

### BulkDeleteFiles.ps1
- **What it does:** Deletes list items in batches from a SharePoint list until the list is empty.
- **Module / platform used:** **SharePoint CSOM** (`ClientContext`, `CamlQuery`, list item deletion APIs).
- **Notable implementation detail:** Uses CAML with a row limit to batch deletes.

### DeleteLibraryScript.ps1
- **What it does:** Deletes a SharePoint library by title.
- **Module / platform used:** **SharePoint CSOM** (`ClientContext`, `SharePointOnlineCredentials`, list deletion APIs).

### CopyFilesBetweenLibraries.ps1
- **What it does:** Recursively copies files and folders from one SharePoint document library to another, skipping the Forms folder and creating destination folders as needed.
- **Module / platform used:** **SharePoint CSOM** (`ClientContext`, folder and file APIs).

### Copy-SPOListViews.ps1
- **What it does:** Copies public views and current-user private views from one SharePoint Online list/library to another and preserves key view properties.
- **Module / platform used:** **SharePoint CSOM** and helper functions for CSOM property loading.
- **Notable implementation detail:** Designed for cross-web / cross-site list view replication.

## 3) SharePoint admin connectivity and tenant operations

### ConnectToAdminPortal.ps1
- **What it does:** Connects to the SharePoint Online admin portal and retrieves deleted containers.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`) plus SharePoint Online admin command usage (`Get-SPODeletedContainer`).

### SetSiteAdmin_MFA_MultipleSites.ps1
- **What it does:** Connects to SharePoint Online and sets site owners / site admins, with CSV path scaffolding for multi-site processing.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`) and **PnP tenant administration** (`Set-PnPTenantSite`).
- **Note:** The inspected content is short and appears to be a partially drafted admin utility.

## 4) Microsoft 365 Groups, licensing, and service principals

### APPClientSecret.ps1
- **What it does:** Enumerates non-Microsoft service principals and retrieves credential metadata for each one, then writes an output report.
- **Module / platform used:** **MSOL / Azure AD PowerShell** (`Connect-MsolService`, `Get-MsolServicePrincipal`, `Get-MsolServicePrincipalCredential`).

### AssignUsersToGroup.ps1
- **What it does:** Retrieves MSOL users matching a display-name filter and contains scaffolding for adding them to a group.
- **Module / platform used:** **MSOL / Azure AD PowerShell** (`Get-MsolUser`, commented `Get-MsolGroup`, commented `Add-MsolGroupMember`).
- **Note:** The add-member action is commented in the inspected content.

### BulkAddGroupMembers.ps1
- **What it does:** Imports usernames from CSV and adds them as members to a Microsoft 365 Group.
- **Module / platform used:** **Exchange Online / Unified Groups PowerShell** (`Add-UnifiedGroupLinks`), with commented remote session import pattern.

## 5) Migration and third-party migration tooling

### BulkSharedDriveMigration.ps1
- **What it does:** Imports migration tasks from CSV and stages file share to SharePoint migration jobs.
- **Module / platform used:** **SharePoint Migration Tool PowerShell** (`Import-Module Microsoft.SharePoint.MigrationTool.PowerShell`, `Register-SPMTMigration`, `Add-SPMTTask`).

### CopyFolder.PS1
- **What it does:** Loads source and target objects and runs a folder migration action between SharePoint locations.
- **Module / platform used:** **Metalogix Content Matrix PowerShell snap-ins** (`Metalogix.System.Commands`, `Metalogix.SharePoint.Migration.Commands`).

## 6) Power Platform / Flow / M365 CLI automation

### GetFlowDefinitionsDefaultEnvironment.ps1
- **What it does:** Finds the default Power Platform environment, lists flows as admin, iterates through the flow inventory, and tracks export progress / missing exported items.
- **Module / platform used:** **Microsoft 365 CLI** (`m365 flow environment list`, `m365 flow list`).
- **Notable implementation detail:** Includes progress reporting and exported-file existence checks.

### SubscribeLibToFlow.ps1
- **What it does:** Connects to SharePoint Online and prepares list/library webhook subscription handling for a large set of list names.
- **Module / platform used:** **PnP PowerShell** (`Connect-PnPOnline`); commented **PnP webhook** command usage (`Add-PnPWebhookSubscription`).
- **Notable implementation detail:** The inspected content shows bulk list-name targeting and webhook inspection scaffolding.

## 7) Microsoft Graph and direct REST-based scripts

### GetSiteUsingmgGraphAPI.ps1
- **What it does:** Obtains a Microsoft Graph access token via client-credentials flow, queries a SharePoint site, and then queries the drives for that site.
- **Module / platform used:** **Microsoft Graph REST API** via `Invoke-RestMethod`.

### createCalendarEventItems.ps1
- **What it does:** Obtains a Microsoft Graph access token and retrieves calendar events for a user.
- **Module / platform used:** **Microsoft Graph REST API** via `Invoke-RestMethod`.

## 8) Certificates and app packaging

### Create-SelfSignedCertificate.ps1
- **What it does:** Creates a self-signed certificate for server-to-server authentication, exports it to PFX, and removes the certificate from the store after export.
- **Module / platform used:** Native **Windows certificate store / PowerShell certificate management**.
- **Notable implementation detail:** The script is built as a reusable parameterized utility with examples and helper functions.

### Create-SelfSignedCertificateNew.ps1
- **What it does:** Uses PowerShell certificate commands to create / export a certificate to PFX and CER formats.
- **Module / platform used:** Native **PowerShell certificate cmdlets** (`New-SelfSignedCertificate`, `Export-PfxCertificate`, `Export-Certificate`).

### createCalendarEventItems_3.ps1
- **What it does:** Generates a Teams app package by copying a manifest template, replacing merge fields, and zipping required assets.
- **Module / platform used:** Native **PowerShell file handling** and `Compress-Archive`.

## 9) Configuration scaffolding / partner pack setup

### Configure-Configs.ps1
- **What it does:** Updates application configuration files for the PnP Partner Pack solution by replacing connection string and app-setting values.
- **Module / platform used:** Native **PowerShell XML / file manipulation** over config files.
- **Notable implementation detail:** Targets multiple `App.config` / `Web.config` files in a PnP Partner Pack solution structure.

---

## Quick module inventory (inspected scripts only)

The following modules / APIs / libraries were directly visible in the inspected script content:

- **PnP PowerShell**
- **SharePoint Online Management Shell**
- **SharePoint CSOM** (`Microsoft.SharePoint.Client` / runtime assemblies)
- **MSOL / Azure AD PowerShell**
- **Exchange Online / Unified Groups PowerShell**
- **SharePoint Migration Tool PowerShell**
- **Microsoft 365 CLI**
- **Microsoft Graph REST API** (`Invoke-RestMethod`)
- **Metalogix Content Matrix PowerShell snap-ins**
- **Native PowerShell certificate cmdlets**
- **Native PowerShell file/XML/archive commands**

---

## Important note

This document is intentionally sanitized:
- No secrets
- No client IDs
- No tenant IDs
- No GUIDs
- No passwords
- No user-specific tokens
- No environment-specific sensitive values

